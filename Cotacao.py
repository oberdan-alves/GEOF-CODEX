# -*- coding: utf-8 -*-
"""
Cotacao.py
==========
MÓDULO DO 1º PASSO — Nova cotação
-----------------------------------
Responsável pelos botões do grupo "1º Passo – Nova cotação":
    - CRIAR PASTA COTAÇÃO
    - PEDIDO (gera o PEDIDO DE COTAÇÃO PDPAS em PDF + JPG)
    - AJUSTE DO PEDIDO (marcar itens como CANCELADO e/ou incluir uma
      OBSERVAÇÃO em destaque, aplicados na próxima geração do PEDIDO)

Também contém os dados fixos (templates) usados apenas nesse documento:
    NUCLEO_TEMPLATES        -> dados de cada núcleo responsável (item 10/11)
    TEXTO_OBSERVACAO_NECFM  -> observação em vermelho exigida quando a
                                origem é NECFM (conserto de equipamentos)

Este módulo importa constantes e utilitários de GeradorCotacao.py (cores,
fontes, formatação, contatos padrão, conversão PDF->JPG) para não repetir
código. A classe CotacaoMixin é "misturada" (herdada) pela classe principal
AppCotacao, então os métodos abaixo usam `self` normalmente para acessar
os campos da tela (self.entry_ano, self.linhas_resultado etc.).

IMPORTANTE sobre PEDIDO x tela "ITENS DA COTAÇÃO": o botão CONSULTAR
COTAÇÃO (GeradorCotacao.py) mostra na tela só os itens da COTAÇÃO ATIVA
(a caixa selecionada na matriz) — self.linhas_resultado reflete sempre
uma única cotação. Já o botão PEDIDO, aqui neste módulo, é independente
disso: ele consulta a aba Controle DIRETAMENTE, um código de cada vez,
para TODAS as cotações preenchidas na matriz (até 10), e gera um PEDIDO
(PDF + JPG) por código, cada um na sua própria pasta
\\COTACAO\\{código}\\ — sem depender do que está sendo exibido na tela
no momento do clique. O mesmo vale para AJUSTE DO PEDIDO
(acao_ajustar_pedido), que consulta direto a cotação ativa em vez de
reaproveitar self.linhas_resultado, para nunca ajustar itens errados
mesmo que a tela esteja mostrando outra coisa.
"""

import os
import json
import tkinter as tk
from Mensagens import messagebox

import customtkinter as ctk
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT

from GeradorCotacao import (
    centralizar_janela,
    COR_FUNDO_APP, COR_TEXTO, COR_DESTAQUE,
    FONTE_ENTRY, FONTE_SECAO, FONTE_BOTAO, FONTE_TABELA,
    COR_CIANO_ADJ, COR_VERDE_ADJ, COR_AMARELO_ADJ,
    CONTATOS_ORDEM_PADRAO,
    converter_pdf_para_jpg, exibir_preview_pdf_jpg, montar_cabecalho_institucional,
    buscar_cotacoes_na_planilha,
)

# ----------------------------------------------------------------------
# MODELOS DO "NÚCLEO" (item 10/10.1/11 do PEDIDO DE COTAÇÃO), por ORIGEM +
# grupo (APS/HRG) — construídos a partir dos 4 modelos reais enviados.
# 'chave_contato' aponta para a MESMA variável configurável em
# CONTATOS_ORDEM_PADRAO, para não duplicar cadastro.
# ----------------------------------------------------------------------
NUCLEO_TEMPLATES = {
    ("NPDOC", "HRG"): {
        "nome_nucleo": "Patrimônio", "sigla": "NPDOC/HRG",
        "horario": "de segunda a QUINTA de 08:00 as 11:00 e 14:00 as 17:00",
        "chave_contato": "NPDOC_HRGINV",
        "cargo": "Chefe do Núcleo de Patrimônio e Documentação Administrativa",
        "extra": " e sala 3449 7083",
    },
    ("NPDOC", "APS"): {
        "nome_nucleo": "Patrimônio", "sigla": "NPDOC/APS",
        "horario": "de 07h às 12h e 14h às 18h, de segunda a quinta-feira",
        "chave_contato": "NPDOC_APSINV",
        "cargo": "Chefe do Núcleo de Patrimônio e Documentação Administrativa",
        "extra": "", "destaque_vermelho": True,
    },
    ("NUAL", "HRG"): {
        "nome_nucleo": "Almoxarifado", "sigla": "NUAL/HRG",
        "horario": "de 08h às 12h e de 14h às 17h, de segunda a sexta-feira",
        "chave_contato": "NUAL_HRG",
        "cargo": "Chefe do Núcleo de Almoxarifado/HRG",
        "extra": "",
    },
    ("NUAL", "APS"): {
        "nome_nucleo": "ALMOXARIFADO da Gerência de Apoio Operacional da Atenção Primária SUL",
        "sigla": "NUAL/GAOAPS", "setor_sigla": "NUAL/GAOAPS-Sul",
        "horario": "de 08h às 12h e de 13h às 16h, de segunda a sexta-feira",
        "chave_contato": "NUAL_APS",
        "cargo": "Chefe do Núcleo de Almoxarifado/GAOAPS-SUL – (61) 98138-6973.",
        "extra": "",
    },
    ("NECFM", "APS"): {
        "nome_nucleo": "ENGENHARIA", "sigla": "NECFM/GAOESP", "setor_sigla": "NECFM/GAOESP-APS",
        "horario": "de 08h às 12h e de 13h às 16h, de segunda a sexta-feira",
        "chave_contato": "NECFM",
        "cargo": "Chefe do Núcleo de Engenharia Clínica e Física Médica/HRG NECFM (MEZANINO ACIMA DO PAI)",
        "extra": "– (61) 9639-0162.", "destaque_vermelho": True, "modelo_entrega": "agendar",
        "destaque_item10": "base",
        "label_item11": "Informações, agendamentos, Marcação de Visitas técnicas e "
                        "responsável pelo Atesto para pagamento:",
    },
    ("NECFM", "HRG"): {
        # modelo próprio (179/2026): diferente da NECFM/APS — NÃO usa a
        # frase "agendar", só o texto padrão de entrega
        "nome_nucleo": "Engenharia Clínica", "sigla": "NECFM/HRG",
        "horario": "de 08h às 12h e de 14h às 17h, de segunda a sexta-feira",
        "chave_contato": "NECFM",
        "cargo": "Chefe do Núcleo de Engenharia Clínica e Física Médica/HRG NECFM (MEZANINO ACIMA DO PAI)",
        "extra": "– (61) 9639-0162.", "destaque_vermelho": True,
        "label_item11": "Informações, agendamentos, Marcação de Visitas técnicas e "
                        "responsável pelo Atesto para pagamento:",
    },
    ("NFL", "APS"): {
        "nome_nucleo": "Farmácia", "sigla": "NFL/APS",
        "horario": "", "chave_contato": "NFL", "cargo": "Farmácia", "extra": "",
    },
    ("NFH", "HRG"): {
        # modelo próprio (153/2026): item 10 com destaque na cor base (ciano)
        "nome_nucleo": "Farmácia", "sigla": "NFH/HRG",
        "horario": "de 08h às 12h e de 14h às 17h, de segunda a sexta-feira",
        "chave_contato": "NFH",
        "cargo": "Chefe do Nucleo de Farmácia/HRG – (61) 3449-7070, 7071.",
        "extra": "", "destaque_item10": "base",
    },
}

# Bloco em vermelho que só entra quando a origem é NECFM (conserto de
# equipamentos exige visita técnica prévia) — texto exatamente do modelo.
TEXTO_OBSERVACAO_NECFM = [
    "- O orçamento deverá ser realizado in loco, mediante visita técnica previamente agendada, "
    "cuja comprovação deverá ser apresentada junto à proposta por meio de Atestado de Visita "
    "Técnica.",
    "- Recomenda-se que os licitantes mantenham contato prévio com a Engenharia Clínica, a fim "
    "de obter informações completas e atualizadas sobre os equipamentos e suas condições de "
    "funcionamento.",
    "- As propostas deverão conter a descrição detalhada dos serviços a serem executados, bem "
    "como a relação das peças a serem substituídas ou reparadas, incluindo especificações "
    "técnicas, prazos de execução e garantias oferecidas.",
    "- A ausência dessas informações ou a apresentação incompleta poderá acarretar a "
    "desclassificação da proposta, em razão da impossibilidade de análise da adequação técnica "
    "e da vantajosidade da contratação.",
]


# ----------------------------------------------------------------------
# GERAÇÃO DO PDF — PEDIDO DE COTAÇÃO PDPAS
# ----------------------------------------------------------------------
def gerar_pdf_pedido_cotacao(caminho_pdf, codigo_cotacao, ano, operador, sei_texto, origem,
                              tipo_planilha, itens, contatos, diretor_nome, diretor_matricula,
                              tr_numero=None):
    """Gera o PEDIDO DE COTAÇÃO PDPAS (documento enviado aos fornecedores
    antes de receber propostas), replicando os 4 modelos reais (APS,
    APS INV, HRG, HRG INV). 'itens' é a lista de itens da cotação ativa
    (já filtrada pelo chamador — ver acao_salvar_cotacao), cada um com
    CODIGO/DESCRICAO/UN/QTD."""
    contatos = contatos or CONTATOS_ORDEM_PADRAO
    tipo_norm = tipo_planilha.strip().upper()
    is_hrg = "HRG" in tipo_norm
    is_inv = "INV" in tipo_norm
    grupo = "HRG" if is_hrg else "APS"
    cor_base = COR_CIANO_ADJ if is_hrg else COR_VERDE_ADJ
    origem_norm = (origem or "").strip().upper()

    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4,
                             topMargin=0.5 * cm, bottomMargin=0.5 * cm,
                             leftMargin=1.2 * cm, rightMargin=1.2 * cm)
    styles = getSampleStyleSheet()
    largura_util = A4[0] - 2.4 * cm

    estilo_publpor = ParagraphStyle("PublPor", parent=styles["Normal"], fontSize=9,
                                    fontName="Helvetica-Bold")
    estilo_barra_esq = ParagraphStyle("BarraEsqPed", parent=styles["Normal"], fontSize=8,
                                      fontName="Helvetica-Bold", alignment=TA_LEFT, leading=9.5)
    estilo_barra_centro = ParagraphStyle("BarraCentroPed", parent=styles["Normal"], fontSize=16, leading=18,
                                         fontName="Helvetica-Bold", alignment=TA_CENTER)
    estilo_barra_direita = ParagraphStyle("BarraDirPed", parent=styles["Normal"], fontSize=10,
                                          fontName="Helvetica-Bold", alignment=TA_RIGHT, leading=12)
    estilo_setor_label = ParagraphStyle("SetorLabel", parent=styles["Normal"], fontSize=8.5, leading=10,
                                        fontName="Helvetica-Bold", alignment=TA_LEFT)
    estilo_setor_valor = ParagraphStyle("SetorValor", parent=styles["Normal"], fontSize=8.5,
                                        leading=10,
                                        fontName="Helvetica-Bold", alignment=TA_CENTER,
                                        textColor=colors.red)
    estilo_prazo = ParagraphStyle("PrazoPed", parent=styles["Normal"], fontSize=9, leading=10.5,
                                  fontName="Helvetica-Bold", alignment=TA_CENTER)
    estilo_entrega_bar = ParagraphStyle("EntregaBar", parent=styles["Normal"], fontSize=9, leading=10.5,
                                        fontName="Helvetica-BoldOblique", alignment=TA_CENTER)
    estilo_corpo = ParagraphStyle("CorpoPed", parent=styles["Normal"], fontSize=8.3, leading=9.3,
                                    alignment=TA_JUSTIFY)
    estilo_corpo_bold = ParagraphStyle("CorpoPedBold", parent=estilo_corpo,
                                       fontName="Helvetica-Bold")
    estilo_corpo_bold_italic = ParagraphStyle("CorpoPedBI", parent=estilo_corpo,
                                              fontName="Helvetica-BoldOblique")
    estilo_corpo_red = ParagraphStyle("CorpoPedRed", parent=estilo_corpo, textColor=colors.red,
                                      fontName="Helvetica-Bold")
    estilo_cel = ParagraphStyle("CelPed", parent=styles["Normal"], fontSize=8, leading=9.7)
    estilo_cel_center = ParagraphStyle("CelPedC", parent=estilo_cel, alignment=TA_CENTER)
    estilo_rodape = ParagraphStyle("RodapePed", parent=styles["Normal"], fontSize=7.3, leading=8.2,
                                   alignment=TA_CENTER, fontName="Helvetica-Bold")

    story = []

    # ---- cabeçalho: brasão do GDF à esquerda + texto institucional
    # centralizado, "Publ. por" / "SEI" logo abaixo, acima da barra ----
    story.append(montar_cabecalho_institucional(largura_util, [
        "GOVERNO DO DISTRITO FEDERAL",
        "SECRETARIA DE ESTADO DE SAÚDE",
        "SUPERINTENDÊNCIA DA REGIÃO DE SAÚDE SUL",
        "DIRETORIA ADMINISTRATIVA",
        f"PROGRAMA DE DESCENTRALIZAÇÃO PROGRESSIVA DE AÇÕES DE SAÚDE - PDPAS"
        f"{'/' + grupo if is_hrg else ''}",
    ]))
    story.append(Spacer(1, 4))

    estilo_publicacao_direita = ParagraphStyle(
        "PublicacaoDireita", parent=estilo_publpor, alignment=TA_RIGHT)
    linha_publ = [
        Paragraph(f"SEI: {sei_texto}", estilo_publpor) if sei_texto else "",
        Paragraph(f"Publ. por: <u>{operador or ''}</u>", estilo_publicacao_direita),
    ]
    tabela_publ = Table([linha_publ], colWidths=[largura_util * 0.6, largura_util * 0.4])
    tabela_publ.setStyle(TableStyle([
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(tabela_publ)
    story.append(Spacer(1, 4))

    # ---- barra de título (3 células) ----
    if is_inv:
        texto_bar1 = "AQUISIÇÃO DE BEM PERMANENTE/INVESTIMENTO"
        cor_bar1 = COR_AMARELO_ADJ
    else:
        texto_bar1 = "SUPERINTENDÊNCIA DA REGIÃO DE SAÚDE SUL -"
        cor_bar1 = colors.white
    texto_bar2 = "HRG" if is_hrg else ("APS Sul" if is_inv else "APS - SUL")

    barra_titulo = Table(
        [[Paragraph(texto_bar1, estilo_barra_esq), Paragraph(texto_bar2, estilo_barra_centro),
          Paragraph(f'PEDIDO DE COTAÇÃO PDPAS &nbsp;&nbsp;'
                    f'<font size="13">{codigo_cotacao}/{ano}</font>', estilo_barra_direita)]],
        colWidths=[largura_util * 0.42, largura_util * 0.16, largura_util * 0.42])
    barra_titulo.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (0, 0), cor_bar1),
        ("BACKGROUND", (1, 0), (1, 0), cor_base),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (0, 0), 8),
    ]))
    story.append(barra_titulo)
    story.append(Spacer(1, 4))

    # ---- Setor Solicitante + prazo de recebimento ----
    sigla_origem = origem_norm
    template_nucleo = NUCLEO_TEMPLATES.get((origem_norm, grupo))
    if template_nucleo:
        sigla_origem = template_nucleo.get("setor_sigla", template_nucleo["sigla"])
    elif origem_norm:
        sigla_origem = f"{origem_norm}/{grupo}"

    caixa_setor = Paragraph(
        f'Setor Solicitante: <font color="red">{sigla_origem}</font>', estilo_setor_label)
    caixa_prazo = Paragraph(
        "2 - Prazo para recebimento das propostas, conforme <nobr>publicado no E-MAIL.</nobr>",
        estilo_prazo)
    linha_setor = Table([[caixa_setor, caixa_prazo]],
                        colWidths=[largura_util * 0.30, largura_util * 0.70])
    linha_setor.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, 0), "CENTER"),
    ]))
    story.append(linha_setor)

    # ---- barra "entrega imediata" (cor base) ----
    barra_entrega = Table([[Paragraph(
        "3 - PREFERENCIALMENTE ENTREGA IMEDIATA APÓS A SOLICITAÇÃO AO FORNECEDOR",
        estilo_entrega_bar)]], colWidths=[largura_util])
    barra_entrega.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), cor_base), ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(barra_entrega)
    story.append(Spacer(1, 4))

    story.append(Paragraph("Prezados Senhores,", estilo_corpo))
    story.append(Spacer(1, 3))
    if is_hrg:
        story.append(Paragraph(
            "1 - A Superintendência da Região de Saúde Sul convoca as empresas interessadas em "
            "participar da DISPENSA DE LICITAÇÃO, que será levada a efeito no Hospital Regional "
            "do Gama, situado na ÁREA ESPECIAL Nº 01 - Setor Central GAMA – Brasília – DF, para "
            "aquisição dos materiais constantes deste Pedido de Cotação PDPAS. Todos os "
            "documentos do processo poderão ser acompanhados pelo site "
            "https://www.portalsei.df.gov.br/", estilo_corpo))
    else:
        story.append(Paragraph(
            "1 - A Superintendência da Região de Saúde Sul convoca as empresas interessadas em "
            "participar da DISPENSA DE LICITAÇÃO, que será levada a efeito na Gerência de "
            "Orçamento e Finanças da Superintendência da Região de Saúde Sul (GEOF/DA/SRSSU/SES) "
            "- Brasília/DF, para aquisição dos materiais e/ou serviços constantes deste Pedido de "
            "Cotação PDPAS. Todos os documentos do processo poderão ser acompanhados pelo site "
            "https://www.portalsei.df.gov.br/", estilo_corpo))
    story.append(Paragraph(
        "A DISPENSA DE LICITAÇÃO será levada a efeito com fundamento no INCISO II, ARTIGO 75, "
        "LEI Nº 14.133/21.", estilo_corpo))
    story.append(Spacer(1, 4))

    # ---- tabela de itens (todos os itens da cotação ativa) ----
    # Um item pode trazer "_observacao_destaque" (ajustado em AJUSTE DO
    # PEDIDO): quando presente, uma linha extra é inserida logo abaixo do
    # item, ocupando a largura toda, em fonte maior/negrito, fundo amarelo
    # e texto vermelho.
    estilo_obs_destaque = ParagraphStyle(
        "ObsDestaquePed", parent=styles["Normal"], fontSize=10, leading=12,
        fontName="Helvetica-Bold", textColor=colors.red, alignment=TA_LEFT)

    dados_tabela = [["ITEM", "CÓD.", "DESCRIÇÃO", "QTD.", "UN"]]
    estilos_extra_itens = []
    linha_idx = 1
    for item in itens:
        qtd = item.get("QTD")
        try:
            qtd_texto = f"{float(qtd):,.3f}".rstrip("0").rstrip(".").replace(",", "X").replace(".", ",").replace("X", ".")
        except (TypeError, ValueError):
            qtd_texto = str(qtd or "")
        dados_tabela.append([
            str(item.get("ITEM") or ""), item.get("CODIGO") or "",
            Paragraph(item.get("DESCRICAO") or "", estilo_cel),
            qtd_texto, item.get("UN") or "",
        ])
        linha_idx += 1

        observacao_destaque = item.get("_observacao_destaque")
        if observacao_destaque:
            # No SPAN do ReportLab só o conteúdo da PRIMEIRA célula do
            # intervalo é desenhado — por isso o texto vai na coluna 0.
            dados_tabela.append([Paragraph(observacao_destaque, estilo_obs_destaque), "", "", "", ""])
            estilos_extra_itens += [
                ("SPAN", (0, linha_idx), (4, linha_idx)),
                ("BACKGROUND", (0, linha_idx), (4, linha_idx), COR_AMARELO_ADJ),
            ]
            linha_idx += 1

    tabela_itens = Table(dados_tabela,
                         colWidths=[1.4 * cm, 1.8 * cm, largura_util - 6.4 * cm, 1.6 * cm, 1.6 * cm],
                         repeatRows=1)
    tabela_itens.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (1, -1), "CENTER"), ("ALIGN", (3, 0), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ] + estilos_extra_itens))
    story.append(tabela_itens)

    # ---- OBSERVAÇÕES ----
    estilo_obs_titulo = ParagraphStyle("ObsTitulo", parent=styles["Normal"], fontSize=9, leading=10,
                                       fontName="Helvetica-Bold", alignment=TA_CENTER)
    story.append(Table([[Paragraph("OBSERVAÇÕES:", estilo_obs_titulo)]], colWidths=[largura_util],
                       style=TableStyle([("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                                         ("TOPPADDING", (0, 0), (-1, -1), 2),
                                         ("BOTTOMPADDING", (0, 0), (-1, -1), 2)])))

    destaque_vermelho = bool(template_nucleo and template_nucleo.get("destaque_vermelho")
                             and origem_norm == "NECFM")
    if destaque_vermelho:
        for linha in TEXTO_OBSERVACAO_NECFM:
            story.append(Paragraph(linha, estilo_corpo_red))

    story.append(Paragraph("4 - O não cumprimento do prazo de entrega ensejará as sanções legais.",
                           estilo_corpo))
    story.append(Paragraph(
        "5 - Normalmente os pagamentos da Região Sul são efetuados por TED ou Transferência, "
        "entre os dias 20 e 30 de cada mês. <b>NÃO EMITIR BOLETO.</b>", estilo_corpo))
    story.append(Paragraph(
        "6 - Haverá retenção de Imposto de Renda no percentual de 1,2% para materiais e 4,8% "
        "para serviços, Art. 2º-A da IN RFB nº 2145 de 26/06/2023.", estilo_corpo))

    if is_inv:
        story.append(Table([[Paragraph(
            "7 - Amostra do item poderá ser solicitada. ENVIAR IMAGENS, FICHA TÉCNICA, FOLDERS.",
            estilo_corpo_bold)]], colWidths=[largura_util],
            style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_base)])))
    else:
        story.append(Paragraph("7 - Amostra do item poderá ser solicitada.", estilo_corpo))

    story.append(Paragraph(
        "8 - O Critério de Julgamento será o MENOR PREÇO e/ou MENOR PRAZO DE ENTREGA.",
        estilo_corpo))
    story.append(Paragraph(
        "9 - A quantidade a ser adquirida poderá ser menor do que a solicitada no Pedido de "
        "Cotação.", estilo_corpo))

    # ---- item 10 / 10.1 / 11: dados do núcleo responsável ----
    if is_hrg:
        cnpj_unidade, insc_ou_endereco = "00.394.700/0006-12", \
            "ÁREA ESPECIAL Nº 01, Lote Único - Setor Central GAMA/DF. CEP: 72.405-901"
    else:
        # nas cotações APS não-INV, o endereço vem acompanhado das coordenadas
        # (Maps / Plus Code); nas APS INV, não aparece.
        cnpj_unidade, insc_ou_endereco = "00.394.700/0051-77", \
            ("Setor Oeste, SOE Q 12, Praça 1,  Área Especial - Gama Oeste/DF, CEP. 72.425-120." +
             (" Ao lado da Igreja Nossa Senhora Aparecida. Coordenadas no Maps "
              "-16.007452594187583, -48.075027885099104" if not is_inv else ""))

    if template_nucleo:
        contato_texto = contatos.get(template_nucleo["chave_contato"], "")
        texto_item10 = (
            f"10 - Entrega/Recolhimento agendar no Núcleo de {template_nucleo['nome_nucleo']} - "
            f"({template_nucleo['sigla']}) {template_nucleo['horario']}."
            if template_nucleo.get("modelo_entrega") == "agendar" else
            f"10 - Entrega no Núcleo de {template_nucleo['nome_nucleo']} "
            f"({template_nucleo['sigla']}) {template_nucleo['horario']}.")

        destaque_item10 = template_nucleo.get("destaque_item10")

        if destaque_item10:
            cor_destaque10 = COR_AMARELO_ADJ if destaque_item10 == "amarelo" else cor_base
            story.append(Table([[Paragraph(texto_item10, estilo_corpo_bold_italic)]],
                               colWidths=[largura_util],
                               style=TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_destaque10)])))
        else:
            story.append(Paragraph(texto_item10, estilo_corpo_bold_italic))

        story.append(Paragraph(
            f"10.1 - CNPJ: {cnpj_unidade}. Endereço: {insc_ou_endereco}", estilo_corpo_bold_italic))
        label_item11 = template_nucleo.get("label_item11", "Informações Técnicas e responsável "
                                                            "pelo Atesto:")
        story.append(Paragraph(
            f"11 - {label_item11} {contato_texto}", estilo_corpo_bold_italic))
        story.append(Paragraph(
            f"{template_nucleo['cargo']}{template_nucleo.get('extra', '')}",
            estilo_corpo_red if destaque_vermelho else estilo_corpo_bold_italic))
    story.append(Spacer(1, 4))

    # ---- item 12 (lista vertical, sem tabela, para ocupar menos espaço —
    # só a linha "12 -" fica em negrito; os sub-itens ficam em fonte normal) ----
    story.append(Paragraph("12 - Solicitamos que as propostas apresentem:", estilo_corpo_bold))
    for texto_12 in [
        "12.1 - Logotipo da empresa timbrado com nome, endereço, CNPJ e telefone;",
        "12.2 - Data, assinatura e dados do responsável pela cotação;",
        "12.3 - Prazo de validade no mínimo 30 dias;",
        "12.4 - Prazo de entrega do item;",
        "12.5 - Especificações do item idênticas ao solicitado no Pedido de Cotação;",
        "12.6 - O valor do frete, imposto e taxas inclusos no preço unitário do item.",
        "12.7 - O preço do medicamento deve estar abaixo do Preço Fábrica da CMED.",
        "12.8 - Dados bancários para pagamento. Nota Fiscal de empresas do DF, acima de "
        "R$ 5.000,00 somente BRB.",
    ]:
        story.append(Paragraph(texto_12, estilo_corpo))

    story.append(Paragraph(
        "13 - Pedimos que as propostas sejam acompanhadas das Certidões Negativas válidas e/ou "
        "SICAF, obrigatórias para a contratação:", estilo_corpo))
    for texto in [
        "13.1 - Certidão Negativa de Débitos de Tributos e Contribuições Federais - SRF-IN nº "
        "80/1997, Art. 68, Inciso III, Lei 14.133/2021;",
        "13.2 - Certificado de Regularidade do FGTS/CEF - Art. 27, da Lei nº 8.036/1990, Art. 68, "
        "Inciso IV, Lei 14.133/2021;",
        "13.3 - Certidão Negativa de Débitos - DF - Art. 56, § único, inciso VI, do Decreto "
        "Distrital nº 16.098/1994, Art. 68, Inciso III, Lei 14.133/2021;",
        "13.4 - Certidão Negativa de Débitos Trabalhistas do TST - Lei nº 12.440/2011, Art. 68, "
        "Inciso V, Lei 14.133/2021;",
        "13.5 - Certidão Negativa Correcional emitida pela CGU - Art. 91, § 4º da Lei 14.133/2021.",
        "13.6 - Para venda de <b>MEDICAMENTOS</b> apresentar <b>AFE</b> para Distribuição de "
        "Medicamentos - RDC 430/2020 ANVISA e <b>Alvará / Licença Sanitária.</b>",
    ]:
        story.append(Paragraph(texto, estilo_corpo))
    story.append(Spacer(1, 4))

    story.append(Paragraph(
        "AUTORIZO a presente Dispensa de Licitação nos termos do inciso II, do Art. 75 da Lei "
        "14.133/21. Caso necessário fica autorizada a republicação.",
        ParagraphStyle("AutorizoPed", parent=estilo_corpo, alignment=TA_CENTER)))
    if diretor_nome:
        story.append(Paragraph(
            f"Diretor(a) Administrativo(a)/SRSSU: {diretor_nome}, Matrícula: {diretor_matricula}",
            ParagraphStyle("DiretorPed", parent=estilo_corpo_bold, alignment=TA_CENTER)))
    story.append(Spacer(1, 0))

    if is_hrg:
        rodape2 = "PDPAS/HRG - GEOF / Diretoria Administrativa / SRSSU / SES"
        rodape3 = "ÁREA ESPECIAL Nº 01, Lote Único - Setor Central GAMA/DF. CEP: 72.405-901"
        rodape4 = "CNPJ: 00.394.700/0006-12 - Fones: (61) 3449-7105"
    else:
        rodape2 = "PDPAS/HRG - GEOF / Diretoria Administrativa / SRSSU / SES"
        rodape3 = "ÁREA ESPECIAL Nº 01, Lote Único - Setor Central GAMA/DF. CEP: 72.405-901"
        rodape4 = "CNPJ: 00.394.700/0006-12 - Fones: (61) 3449-7105"
    story.append(Paragraph(
        "Brasília – Patrimônio Cultural da Humanidade<br/>" + rodape2 + "<br/>" + rodape3 +
        "<br/>" + rodape4, estilo_rodape))

    doc.build(story)


# ----------------------------------------------------------------------
# AJUSTE DO PEDIDO — persistência e aplicação de cancelamentos/observações
# ----------------------------------------------------------------------
def _caminho_arquivo_ajustes_pedido(pasta_cotacao, codigo_cotacao):
    return os.path.join(pasta_cotacao, f"AJUSTES_PEDIDO_{codigo_cotacao}.json")


def _carregar_ajustes_pedido(pasta_cotacao, codigo_cotacao):
    """Lê o arquivo AJUSTES_PEDIDO_{cotação}.json salvo pela tela AJUSTE DO
    PEDIDO. Devolve um dict {numero_item_str: {"cancelado": bool,
    "observacao": str}} — vazio se não houver nada salvo ainda."""
    caminho = _caminho_arquivo_ajustes_pedido(pasta_cotacao, codigo_cotacao)
    if not os.path.exists(caminho):
        return {}
    try:
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f).get("itens", {})
    except Exception:
        return {}


def _aplicar_ajustes_pedido(itens, ajustes):
    """Retorna uma CÓPIA da lista de itens com os ajustes (CANCELAR /
    OBSERVAÇÃO) feitos em AJUSTE DO PEDIDO já aplicados:
      - item cancelado -> DESCRIÇÃO vira "CANCELADO" (mais a observação,
        se houver);
      - item com observação mas NÃO cancelado -> a descrição original é
        mantida, e a observação fica marcada em "_observacao_destaque"
        para ser desenhada em destaque (fonte maior, negrito, fundo
        amarelo, texto vermelho) logo abaixo dela no PDF.
    Não altera a lista original (self.linhas_resultado), que continua
    intacta para uso nos demais passos (Mapa Comparativo, Ordem etc.)."""
    if not ajustes:
        return itens
    resultado = []
    for item in itens:
        item_copia = dict(item)
        ajuste = ajustes.get(str(item_copia.get("ITEM")))
        if ajuste:
            observacao = (ajuste.get("observacao") or "").strip()
            if ajuste.get("cancelado"):
                nova_descricao = "<b>CANCELADO</b>"
                if observacao:
                    nova_descricao += f" — {observacao}"
                item_copia["DESCRICAO"] = nova_descricao
            elif observacao:
                item_copia["_observacao_destaque"] = observacao
        resultado.append(item_copia)
    return resultado


# ----------------------------------------------------------------------
# MIXIN — métodos de ação ligados aos botões do 1º Passo
# ----------------------------------------------------------------------
class CotacaoMixin:
    """Métodos herdados por AppCotacao. Esperam encontrar em `self` os
    widgets e dados montados em GeradorCotacao.py (self.entry_pasta_geof,
    self.entry_ano, self.linhas_resultado, self._sufixo_planilha_atual etc.)."""

    def _obter_pasta_cotacao(self, codigo_cotacao, incluir_propostas=False):
        """Valida PASTA DE TRABALHO / ANO / NOME PLANILHA e devolve
        (pasta, codigo_cotacao, ano) para o código informado, criando a
        pasta se ainda não existir. Devolve None (já mostrando o aviso ou
        erro correspondente) se algo estiver faltando ou codigo_cotacao
        vier vazio. Diferente de _obter_pasta_cotacao_ativa, este NÃO
        depende da caixa ativa na matriz — serve para montar a pasta de
        QUALQUER código (usado pelo PEDIDO, que gera um PDF por cotação
        preenchida, uma de cada vez).

        Com incluir_propostas=True, devolve o caminho até a subpasta
        PROPOSTAS (…\\COTACAO\\{código}\\PROPOSTAS); caso contrário, devolve
        o caminho até a pasta da cotação (…\\COTACAO\\{código})."""
        pasta_geof = self.entry_pasta_geof.get().strip()
        ano = self.entry_ano.get().strip()
        tipo = getattr(self, "_sufixo_planilha_atual", "")
        if not pasta_geof or not ano or not tipo or not codigo_cotacao:
            messagebox.showwarning(
                "Atenção", "Confira PASTA DE TRABALHO, ANO, NOME PLANILHA e a COTAÇÃO.")
            return None

        partes = [pasta_geof, tipo, f"PDPAS {ano}", "COTACAO", codigo_cotacao]
        if incluir_propostas:
            partes.append("PROPOSTAS")
        pasta = os.path.join(*partes)
        try:
            os.makedirs(pasta, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Erro ao criar pasta", str(e))
            return None

        return pasta, codigo_cotacao, ano

    def _obter_pasta_cotacao_ativa(self, incluir_propostas=False):
        """Mesma coisa que _obter_pasta_cotacao, mas usando o código da
        COTAÇÃO ATIVA (a caixa selecionada na matriz). Usado por AJUSTE DO
        PEDIDO (aqui) e por ACESSAR PROPOSTAS (Mapa.py)."""
        codigo_cotacao = (self.entradas_cotacao[self.indice_cotacao_ativa].get().strip()
                           if self.indice_cotacao_ativa is not None else "")
        if not codigo_cotacao:
            messagebox.showwarning(
                "Atenção", "Clique numa das caixas de COTAÇÃO para defini-la como ativa.")
            return None
        return self._obter_pasta_cotacao(codigo_cotacao, incluir_propostas=incluir_propostas)

    def acao_criar_pasta_cotacao(self):
        pasta_geof = self.entry_pasta_geof.get().strip()
        ano = self.entry_ano.get().strip()
        tipo = getattr(self, "_sufixo_planilha_atual", "")

        if not pasta_geof:
            messagebox.showwarning("Atenção", "Informe a PASTA DE TRABALHO (caminho até \\GEOF).")
            return
        if not ano or len(ano) != 4:
            messagebox.showwarning("Atenção", "Informe o ANO com 4 dígitos.")
            return
        if not tipo:
            messagebox.showwarning("Atenção", "Escolha o NOME PLANILHA (define a pasta APS/HRG/...).")
            return

        codigos = [e.get().strip() for e in self.entradas_cotacao if e.get().strip()]
        codigos_invalidos = [c for c in codigos if len(c) != 3]
        if codigos_invalidos:
            messagebox.showwarning(
                "Atenção",
                "Todas as cotações preenchidas precisam ter o código completo (3 dígitos).\n"
                f"Confira: {', '.join(codigos_invalidos)}")
            return
        if not codigos:
            messagebox.showwarning(
                "Atenção", "Digite ao menos uma COTAÇÃO (código de 3 dígitos) para criar a pasta.")
            return

        pasta_base_cotacoes = os.path.join(pasta_geof, tipo, f"PDPAS {ano}", "COTACAO")

        criadas, ja_existiam, erros = [], [], []
        for codigo in codigos:
            pasta_cotacao = os.path.join(pasta_base_cotacoes, codigo)
            pasta_propostas = os.path.join(pasta_cotacao, "PROPOSTAS")
            try:
                existia = os.path.isdir(pasta_cotacao) and os.path.isdir(pasta_propostas)
                os.makedirs(pasta_propostas, exist_ok=True)
                (ja_existiam if existia else criadas).append(codigo)
            except Exception as e:
                erros.append(f"{codigo}: {e}")

        partes = []
        if criadas:
            partes.append("Pastas criadas para a(s) cotação(ões): " + ", ".join(criadas))
        if ja_existiam:
            partes.append("Já existiam (nenhuma alteração): " + ", ".join(ja_existiam))
        if erros:
            partes.append("Erros:\n" + "\n".join(erros))

        titulo = "Pastas de cotação" if not erros else "Pastas de cotação (com erros)"
        messagebox.showinfo(titulo, f"Local: {pasta_base_cotacoes}\n\n" + "\n\n".join(partes))
        self._salvar_defaults_atuais()

    def acao_salvar_cotacao(self):
        """Botão PEDIDO: gera o PEDIDO DE COTAÇÃO PDPAS (PDF + imagem) de
        CADA cotação preenchida na matriz (até 10 caixas) — uma consulta e
        um PDF por código, cada um salvo na sua própria pasta
        \\COTACAO\\{código}\\. Não depende do que está sendo exibido na
        tela "ITENS DA COTAÇÃO" (que mostra só a cotação ativa): para cada
        código preenchido, este método consulta a aba Controle direto e
        gera o PEDIDO correspondente. Havendo só uma cotação preenchida,
        gera normalmente só essa; havendo várias, gera todas de uma vez."""
        self._salvar_defaults_atuais()

        ano = self.entry_ano.get().strip()
        if not ano or len(ano) != 4:
            messagebox.showwarning("Atenção", "Informe o ANO com 4 dígitos.")
            return

        caminho_arquivo = self._obter_caminho_arquivo()
        if not caminho_arquivo:
            messagebox.showwarning(
                "Atenção",
                "Não foi possível localizar o arquivo Excel.\n"
                "Confira a PASTA (até \\GEOF) e clique em CONSULTAR MATRIX.")
            return

        codigos = [e.get().strip() for e in self.entradas_cotacao if e.get().strip()]
        codigos_invalidos = [c for c in codigos if len(c) != 3]
        if codigos_invalidos:
            messagebox.showwarning(
                "Atenção",
                "Todas as cotações preenchidas precisam ter o código completo (3 dígitos).\n"
                f"Confira: {', '.join(codigos_invalidos)}")
            return
        if not codigos:
            messagebox.showwarning(
                "Atenção", "Digite ao menos uma COTAÇÃO (código de 3 dígitos) para gerar o PEDIDO.")
            return

        tipo = getattr(self, "_sufixo_planilha_atual", "")
        operador = self.var_operador.get()
        diretor_nome = self.entry_da_nome.get().strip()
        diretor_matricula = self.entry_da_matricula.get().strip()
        contatos = self.config_dados.get("contatos_ordem", CONTATOS_ORDEM_PADRAO)

        gerados, sem_itens, erros = [], [], []

        for codigo in codigos:
            try:
                itens_do_codigo = buscar_cotacoes_na_planilha(caminho_arquivo, [codigo], ano)
            except Exception as e:
                erros.append(f"{codigo}: erro ao consultar — {e}")
                continue
            if not itens_do_codigo:
                sem_itens.append(codigo)
                continue

            contexto = self._obter_pasta_cotacao(codigo)
            if contexto is None:
                erros.append(f"{codigo}: não foi possível preparar a pasta da cotação.")
                continue
            pasta_cotacao, codigo_cotacao, _ano = contexto

            origem = (itens_do_codigo[0].get("ORIGEM") or "").strip()
            sei_texto = itens_do_codigo[0].get("SEI") or ""

            ajustes = _carregar_ajustes_pedido(pasta_cotacao, codigo_cotacao)
            itens_para_pdf = _aplicar_ajustes_pedido(itens_do_codigo, ajustes)

            nome_arquivo = f"PEDIDO DE COTAÇÃO {codigo_cotacao} {tipo}({origem})"
            caminho_pdf = os.path.join(pasta_cotacao, nome_arquivo + ".pdf")
            caminho_jpg_base = os.path.join(pasta_cotacao, nome_arquivo)

            try:
                gerar_pdf_pedido_cotacao(
                    caminho_pdf, codigo_cotacao, ano, operador, sei_texto, origem, tipo,
                    itens_para_pdf, contatos, diretor_nome, diretor_matricula)
                caminho_jpg = converter_pdf_para_jpg(caminho_pdf, caminho_jpg_base)
                gerados.append((codigo_cotacao, caminho_pdf, caminho_jpg))
            except Exception as e:
                erros.append(f"{codigo}: erro ao gerar PDF — {e}")

        if not gerados:
            partes_erro = []
            if sem_itens:
                partes_erro.append("Sem itens na aba Controle: " + ", ".join(sem_itens))
            if erros:
                partes_erro.append("Erros:\n" + "\n".join(erros))
            messagebox.showwarning(
                "Nenhum PEDIDO gerado",
                "Não foi possível gerar nenhum PEDIDO.\n\n" + "\n\n".join(partes_erro))
            return

        resumo = "\n".join(f"• Cotação {cod} → {os.path.basename(pdf)}"
                           for cod, pdf, _jpg in gerados)
        partes_finais = [f"{len(gerados)} PEDIDO(S) gerado(s):\n{resumo}"]
        if sem_itens:
            partes_finais.append(
                "Sem itens na aba Controle (não gerados): " + ", ".join(sem_itens))
        if erros:
            partes_finais.append("Erros:\n" + "\n".join(erros))
        messagebox.showinfo("Pedido(s) de Cotação gerado(s)", "\n\n".join(partes_finais))

        # pré-visualiza o primeiro PEDIDO gerado
        _codigo_gerado, _pdf_gerado, caminho_jpg = gerados[0]
        exibir_preview_pdf_jpg(self, caminho_jpg, "Pedido de Cotação")

    # ==================================================================
    # AJUSTE DO PEDIDO — cancelar itens e/ou incluir observação em destaque
    # ==================================================================
    def acao_ajustar_pedido(self):
        """Abre uma tela listando os itens da cotação ATIVA para marcar
        quais devem aparecer como CANCELADO no PEDIDO (a descrição do item
        vira "CANCELADO" + a observação, se houver) e/ou incluir uma
        OBSERVAÇÃO que aparece em destaque (fonte maior, negrito, fundo
        amarelo, texto vermelho) logo abaixo da descrição — só quando o
        item NÃO estiver marcado como cancelado. Os ajustes ficam salvos
        por cotação e são aplicados automaticamente na próxima vez que o
        botão PEDIDO for clicado para essa mesma cotação.

        Consulta a aba Controle DIRETO pelo código da cotação ativa (em
        vez de reaproveitar self.linhas_resultado), para nunca ajustar os
        itens errados mesmo que a tela esteja mostrando outra cotação."""
        if self.indice_cotacao_ativa is None:
            messagebox.showwarning(
                "Atenção", "Clique numa das caixas de COTAÇÃO para defini-la como ativa.")
            return
        codigo_ativo = self.entradas_cotacao[self.indice_cotacao_ativa].get().strip()
        if len(codigo_ativo) != 3:
            messagebox.showwarning(
                "Atenção", "A COTAÇÃO ativa precisa ter o código completo (3 dígitos).")
            return

        ano = self.entry_ano.get().strip()
        if not ano or len(ano) != 4:
            messagebox.showwarning("Atenção", "Informe o ANO com 4 dígitos.")
            return

        caminho_arquivo = self._obter_caminho_arquivo()
        if not caminho_arquivo:
            messagebox.showwarning(
                "Atenção",
                "Não foi possível localizar o arquivo Excel.\n"
                "Confira a PASTA (até \\GEOF) e clique em CONSULTAR MATRIX.")
            return

        try:
            itens_da_cotacao_ativa = buscar_cotacoes_na_planilha(caminho_arquivo, [codigo_ativo], ano)
        except Exception as e:
            messagebox.showerror("Erro ao consultar", str(e))
            return
        if not itens_da_cotacao_ativa:
            messagebox.showinfo(
                "Sem resultados",
                f"Nenhum item encontrado na aba 'Controle' para a cotação {codigo_ativo}.")
            return

        contexto = self._obter_pasta_cotacao_ativa()
        if contexto is None:
            return
        pasta_cotacao, codigo_cotacao, _ano = contexto

        ajustes_existentes = _carregar_ajustes_pedido(pasta_cotacao, codigo_cotacao)
        self._abrir_janela_ajuste_pedido(
            pasta_cotacao, codigo_cotacao, itens_da_cotacao_ativa, ajustes_existentes)

    def _abrir_janela_ajuste_pedido(self, pasta_cotacao, codigo_cotacao, itens, ajustes_existentes):
        janela = ctk.CTkToplevel(self)
        janela.title(f"Ajuste do Pedido - Cotação {codigo_cotacao}")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("1150x700")
        centralizar_janela(janela)
        janela.grab_set()

        ctk.CTkLabel(
            janela,
            text="Marque CANCELAR para os itens que não devem ser cotados — a descrição vira "
                 "\"CANCELADO\" (mais a observação, se preenchida). Preencha OBSERVAÇÃO sem marcar "
                 "CANCELAR para incluir uma nota em destaque logo abaixo da descrição do item no "
                 "PEDIDO.",
            font=FONTE_SECAO, text_color=COR_TEXTO, wraplength=1100, justify="left"
        ).pack(anchor="w", padx=16, pady=(16, 8))

        cabecalho = ctk.CTkFrame(janela, fg_color="#1b1b1b")
        cabecalho.pack(fill="x", padx=16)
        titulos = ["ITEM", "CÓDIGO", "DESCRIÇÃO", "CANCELAR", "OBSERVAÇÃO"]
        larguras = [50, 90, 380, 90, 380]
        for i, (t, w) in enumerate(zip(titulos, larguras)):
            ctk.CTkLabel(cabecalho, text=t, font=("Segoe UI", 11, "bold"),
                         text_color=COR_TEXTO, width=w).grid(row=0, column=i, padx=3, pady=4)

        scroll = ctk.CTkScrollableFrame(janela, fg_color="#161616", height=480)
        scroll.pack(fill="both", expand=True, padx=16, pady=(4, 8))

        entradas = []
        for item in itens:
            numero = item.get("ITEM")
            codigo = str(item.get("CODIGO") or "")
            descricao = str(item.get("DESCRICAO") or "")
            salvo = ajustes_existentes.get(str(numero), {})

            linha = ctk.CTkFrame(scroll, fg_color="transparent")
            linha.pack(fill="x", pady=3)

            ctk.CTkLabel(linha, text=str(numero or ""), font=FONTE_TABELA, width=50
                         ).grid(row=0, column=0, padx=3)
            ctk.CTkLabel(linha, text=codigo, font=FONTE_TABELA, width=90
                         ).grid(row=0, column=1, padx=3)
            ctk.CTkLabel(linha, text=descricao[:75], font=FONTE_TABELA, width=380,
                         anchor="w", justify="left").grid(row=0, column=2, padx=3, sticky="w")

            var_cancelar = tk.BooleanVar(value=bool(salvo.get("cancelado", False)))
            ctk.CTkCheckBox(linha, text="", variable=var_cancelar, width=90
                            ).grid(row=0, column=3, padx=3)

            entry_obs = ctk.CTkEntry(linha, font=FONTE_TABELA, width=380)
            entry_obs.grid(row=0, column=4, padx=3)
            entry_obs.insert(0, salvo.get("observacao", ""))

            entradas.append({"numero_item": numero, "var_cancelar": var_cancelar,
                             "entry_obs": entry_obs})

        def salvar():
            ajustes = {}
            for reg in entradas:
                cancelado = reg["var_cancelar"].get()
                observacao = reg["entry_obs"].get().strip()
                if cancelado or observacao:
                    ajustes[str(reg["numero_item"])] = {
                        "cancelado": cancelado, "observacao": observacao}
            conteudo = {"cotacao": codigo_cotacao, "itens": ajustes}
            try:
                with open(_caminho_arquivo_ajustes_pedido(pasta_cotacao, codigo_cotacao),
                          "w", encoding="utf-8") as f:
                    json.dump(conteudo, f, ensure_ascii=False, indent=2)
            except Exception as e:
                messagebox.showerror("Erro ao salvar", f"Não foi possível salvar os ajustes:\n{e}")
                return
            messagebox.showinfo(
                "Ajustes salvos",
                "Os ajustes do Pedido foram salvos.\nClique em PEDIDO novamente para gerar o "
                "documento já com essas alterações.")
            janela.destroy()

        rodape = ctk.CTkFrame(janela, fg_color="transparent")
        rodape.pack(pady=(0, 16))
        ctk.CTkButton(rodape, text="Salvar", font=FONTE_BOTAO, height=36, fg_color=COR_DESTAQUE,
                      command=salvar).pack(side="left", padx=6)
        ctk.CTkButton(rodape, text="Cancelar", font=FONTE_BOTAO, height=36, fg_color="#7a2626",
                      command=janela.destroy).pack(side="left", padx=6)