# -*- coding: utf-8 -*-
"""
Emails.py
=========
MÓDULO DE E-MAILS
------------------
Responsável pelos botões do grupo "E-mails":
    - E-mail Cotação  (monta ASSUNTO + CORPO do e-mail de Pedido de
                        Cotação, prontos para copiar, a partir das
                        cotações ativas na matriz da tela principal)
    - E-mail Ordem    (monta ASSUNTO + CORPO do e-mail de Ordem de Compra,
                        um por fornecedor vencedor, a partir das decisões
                        já tomadas em ADJUDICAÇÃO)
    - FERIADOS        (cadastro de datas de feriado, usado para calcular o
                        prazo de recebimento das propostas em dias úteis)

Fluxo do E-mail Cotação:
    1) Lê todas as cotações preenchidas na matriz (até 10 caixas) e busca
       os itens de cada uma na aba "Controle" da planilha Matrix ativa.
    2) Para cada cotação, monta o OBJETO automático (reaproveitando
       gerar_objeto_automatico, função compartilhada em GeradorCotacao.py)
       e lê o SEI.
    3) Pergunta ao usuário: PUBLICAÇÃO (1ª a 10ª), PRAZO em dias úteis
       (padrão 3), HORAS de recebimento (HH:MM) e se é com EXCLUSIVIDADE.
    4) Calcula DATARECEBIMENTO somando o PRAZO em dias úteis a partir de
       hoje, pulando sábados, domingos e os feriados cadastrados
       (FERIADOS, ver acao_gerenciar_feriados).
    5) Monta o ASSUNTO e o CORPO do e-mail (texto padrão da SES/PDPAS,
       igual aos modelos reais) e exibe numa tela com botões de Copiar.

Fluxo do E-mail Ordem:
    1) Reaproveita _montar_dados_ordens_ativas() (AdjudicacaoOrdem.py) —
       o mesmo agrupamento de itens vencidos por fornecedor usado para
       gerar os PDFs de Ordem de Compra.
    2) Pergunta o nome/matrícula/cargo de quem vai assinar os e-mails
       (flexível — pode ser qualquer operador, não fixo).
    3) Pergunta, opcionalmente, uma OBSERVAÇÃO por fornecedor (ex.: pedir
       aprovação de marca/modelo antes do faturamento).
    4) Monta um ASSUNTO + CORPO por fornecedor vencedor, replicando os
       modelos reais da SES/PDPAS: cabeçalho colorido por tipo (APS=verde
       claro, APS INV=verde escuro, HRG=azul claro, HRG INV=azul escuro),
       com o caso especial da NECFM/Engenharia Clínica (SERVIÇO DE
       MANUTENÇÃO, sempre em vermelho) substituindo o cabeçalho normal; e
       CONTATOS ÚTEIS específicos por SUFIXO_PLANILHA/origem.
    5) Exibe os e-mails gerados numa tela com navegação por fornecedor
       (Fornecedor N de M) e botões de Copiar.

O cadastro de feriados é persistido em feriados.json (funções
carregar_feriados/salvar_feriados em GeradorCotacao.py) e reaproveitado
sempre que um novo e-mail de cotação for montado.
"""

import datetime
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from GeradorCotacao import (
    COR_FUNDO_APP, COR_TEXTO, COR_TEXTO_SEC, COR_DESTAQUE,
    FONTE_ENTRY, FONTE_SECAO, FONTE_BOTAO, FONTE_TABELA,
    formatar_data, formatar_hora, formatar_apenas_digitos,
    carregar_feriados, salvar_feriados, obter_datas_feriados, somar_dias_uteis,
    buscar_cotacoes_na_planilha, gerar_objeto_automatico, CONTATOS_ORDEM_PADRAO,
    salvar_config,
)

# ----------------------------------------------------------------------
# OPÇÕES DE PUBLICAÇÃO (1ª a 10ª)
# ----------------------------------------------------------------------
PUBLICACOES_DISPONIVEIS = [f"{n}ª PUBLICAÇÃO" for n in range(1, 11)]

# ----------------------------------------------------------------------
# TEXTO PADRÃO DO CORPO DO E-MAIL (itens 3 a 13.5 — igual para todas as
# cotações/publicações; só os itens 1, 2 e a lista de cotações variam)
# ----------------------------------------------------------------------
TEXTO_CORPO_BASE = [
    "3 - PREFERENCIALMENTE PRAZO DE ENTREGA IMEDIATO.",
    "",
    "4 - O não cumprimento do prazo de entrega ensejará as sanções legais.",
    "5 - Os pagamentos da Região Sul são efetuados por TED ou Transferência, entre os dias 20 "
    "e 30 de cada mês.",
    "NÃO EMITIR BOLETO.",
    "6 - Haverá retenção de Imposto de Renda no percentual de 1,2% para materiais e 4,8% para "
    "serviços, IN RFB nº 2145 de 26/06/2023.",
    "7 - Amostra do item poderá ser solicitada.",
    "8 - O Critério de Julgamento será o MENOR PREÇO e/ou MENOR PRAZO DE ENTREGA.",
    "9 - A quantidade a ser adquirida poderá ser menor do que a solicitada no Pedido de Cotação.",
    "10 - Os dados de faturamento e local de entrega serão informados no Pedido de Cotação e "
    "Ordem de Compra.",
    "11 - Informações técnicas devem ser sanadas com o contato informado no Pedido de Cotação.",
    "12 - Solicitamos que as propostas apresentem:",
    "    12.1 - Logotipo da empresa timbrado com nome, endereço, CNPJ e telefone;",
    "    12.2 - Data, assinatura e dados do responsável pela cotação;",
    "    12.3 - Prazo de validade no mínimo 30 dias;",
    "    12.4 - Prazo de entrega do item;",
    "    12.5 - Especificações do item idênticas ao solicitado no Pedido de Cotação;",
    "    12.6 - O valor do frete, imposto e taxas inclusos no preço unitário do item.",
    "    12.7 - O preço do medicamento deve estar abaixo do Preço Fábrica da CMED.",
    "    12.8 - Dados bancários para pagamento. Nota Fiscal de empresas do DF, acima de "
    "R$ 5.000,00 somente BRB.",
    "13 - Pedimos que as propostas sejam acompanhadas das Certidões Negativas válidas, "
    "obrigatórias para a contratação:",
    "    13.1 - Certidão Negativa de Débitos de Tributos e Contribuições Federais - SRF-IN nº "
    "80/1997;",
    "    13.2 - Certificado de Regularidade do FGTS/CEF - Art. 27, da Lei nº 8.036/1990;",
    "    13.3 - Certidão Negativa de Débitos - DF - Art. 56, § único, inciso VI, do Decreto "
    "Distrital nº 16.098/1994;",
    "    13.4 - Certidão Negativa de Débitos Trabalhistas do TST - Lei nº 12.440/2011;",
    "    13.5 - Certidão Negativa Correcional emitida pela CGU - Art. 91, § 4º da Lei "
    "14.133/2021",
]

# ----------------------------------------------------------------------
# TEXTO SOBRE CARTA DE EXCLUSIVIDADE (só entra quando o usuário marcar a
# cotação como sendo com EXCLUSIVIDADE)
# ----------------------------------------------------------------------
TEXTO_EXCLUSIVIDADE = [
    "SOBRE CARTA DE EXCLUSIVIDADE:",
    "Nova Lei de Licitações 14.133:",
    "Art. 74. É inexigível a licitação quando inviável a competição, em especial nos casos de: "
    "§ 3º Para fins do disposto no inciso III do caput deste artigo, considera-se de notória "
    "especialização o profissional ou a empresa cujo conceito no campo de sua especialidade, "
    "decorrente de desempenho anterior, estudos, experiência, publicações, organização, "
    "aparelhamento, equipe técnica ou outros requisitos relacionados com suas atividades, "
    "permita inferir que o seu trabalho é essencial e reconhecidamente adequado à plena "
    "satisfação do objeto do contrato.",
    "Art. 75. É dispensável a licitação:",
    "     III - para contratação que mantenha todas as condições definidas em edital de "
    "licitação realizada há menos de 1 (um) ano, quando se verificar que naquela licitação:",
    "    a) não surgiram licitantes interessados ou não foram apresentadas propostas válidas;",
    "    b) as propostas apresentadas consignarem preços manifestamente superiores aos "
    "praticados no mercado ou incompatíveis com os fixados pelos órgãos oficiais competentes;",
    "Decreto 44330 de 16/03/2023 Art. 88. No processo licitatório para aquisição de bens e "
    "contratação de serviços em geral, o valor estimado será definido com base no melhor preço "
    "aferido por meio da utilização dos seguintes parâmetros, adotados de forma combinada ou "
    "não:",
    "    I - relatório de pesquisa de preços de produtos com base nas informações da Nota "
    "Fiscal eletrônica – Nfe;",
    "    II - preços públicos referentes a aquisições ou contratações similares realizadas "
    "pelo Distrito Federal e demais entes públicos;",
    "    III - pesquisa publicada em mídias ou sítios especializados ou de domínio amplo.",
    "    IV - pesquisa direta com, no mínimo, três fornecedores, mediante solicitação formal "
    "de cotação, por meio de carta ou e-mail",
    "Art. 105. Nas contratações diretas por inexigibilidade ou por dispensa de licitação, "
    "aplica-se o disposto no art. 90.",
    "    § 1º Quando não for possível estimar o valor do objeto na forma estabelecida no art. "
    "90, a justificativa de preços será dada com base em valores de contratações de objetos "
    "idênticos, comercializados pela futura contratada, por meio da apresentação de notas "
    "fiscais emitidas para outros contratantes, públicos ou privados, no período de até 1 "
    "(um) ano anterior à data da contratação pela Administração, ou por outro meio idôneo.",
    "    § 2º Excepcionalmente, caso a futura contratada não tenha comercializado o objeto "
    "anteriormente, a justificativa de preço de que trata o parágrafo anterior poderá ser "
    "realizada com objetos semelhantes de mesma natureza, devendo apresentar especificações "
    "técnicas que demonstrem similaridade com o objeto pretendido.",
    "    § 3º Fica vedada a contratação direta por inexigibilidade caso a justificativa de "
    "preços demonstra a possibilidade de competição.",
]


# ----------------------------------------------------------------------
# E-MAIL ORDEM — cores do cabeçalho por tipo de planilha
# ----------------------------------------------------------------------
# Ver os modelos reais de referência: APS = verde claro, APS INV = verde
# escuro, HRG = azul claro, HRG INV = azul escuro. A NECFM (Engenharia
# Clínica — cotações de SERVIÇO DE MANUTENÇÃO) sempre usa vermelho,
# independentemente do tipo de planilha — ela substitui completamente o
# cabeçalho normal (ver montar_texto_email_ordem).
COR_CABECALHO_ORDEM_APS = "#22C55E"        # verde claro
COR_CABECALHO_ORDEM_APS_INV = "#15803D"    # verde escuro
COR_CABECALHO_ORDEM_HRG = "#3B82F6"        # azul claro
COR_CABECALHO_ORDEM_HRG_INV = "#1E3A8A"    # azul escuro
COR_CABECALHO_ORDEM_NECFM = "#DC2626"      # vermelho (Engenharia Clínica)
COR_DESTAQUE_OBSERVACAO_ORDEM = "#DC2626"  # vermelho, mesma cor da NECFM

# Contato institucional fixo que aparece em TODOS os e-mails de Ordem,
# independentemente do tipo de planilha ou da origem (ver os 5 modelos de
# referência — essa linha é idêntica em todos eles).
CONTATO_PDPAS_PADRAO_ORDEM = "Oberdan Alves – GEOF/SRSSU – (61) 3449-7105 – Preferencialmente por e-mail"

TEXTO_ENCERRAMENTO_EXERCICIO_ORDEM = (
    "    Em atenção à Portaria SES-DF nº 473/2023 (PDPAS) e à Lei nº 14.133/2021, art. 137, "
    "§1º, informamos que entregas ou faturamentos não concluídos até o fechamento do "
    "exercício orçamentário serão automaticamente cancelados, sem geração de novo empenho "
    "para o exercício seguinte."
)
TEXTO_SANCOES_ORDEM = (
    "    O atraso injustificado ou o descumprimento das condições pactuadas poderá ensejar "
    "sanções administrativas, nos termos dos arts. 155 a 159 da Lei nº 14.133/2021, incluindo "
    "advertência, multa e impedimento de licitar e contratar com a Administração Pública."
)


def _dados_institucionais_email_ordem(is_hrg):
    """Dados da unidade recebedora (usados no item 'Emissão da Nota
    Fiscal') — os mesmos valores já usados no PDF da Ordem de Compra
    (ver gerar_pdf_ordem_compra em AdjudicacaoOrdem.py), para manter os
    dois documentos consistentes entre si."""
    if is_hrg:
        return {
            "nome_unidade": "HOSPITAL REGIONAL DO GAMA - DF",
            "cnpj": "00.394.700/0006-12",
            "insc_municipal": "0732950500936",
            "endereco": "ÁREA ESPECIAL Nº 01, Lote Único - Setor Central GAMA/DF. CEP: 72.405-901",
        }
    return {
        "nome_unidade": "SUPERINTENDÊNCIA DA REGIÃO DE SAÚDE SUL – APS",
        "cnpj": "00.394.700/0051-77",
        "insc_municipal": "0732950502475",
        "endereco": "GAOAPS-SUL, Setor Oeste, SOE Q 12, Praça 1, Área Especial – Gama Oeste/DF, "
                    "CEP 72.425-120",
    }


def _contatos_uteis_email_ordem(is_hrg, is_inv, is_necfm, contatos):
    """Monta a lista de linhas (assunto, contato) da tabela CONTATOS ÚTEIS,
    que muda conforme o SUFIXO_PLANILHA (APS/APS INV/HRG/HRG INV) — e,
    quando a origem é NECFM (Engenharia Clínica), usa um conjunto próprio
    de contatos independentemente do tipo de planilha, igual ao modelo
    real de referência (601/2025-APS - CONSERTO CAMARA CONSERVAÇÃO)."""
    linhas = [("Dúvidas sobre o processo de dispensa e documentação (PDPAS)",
               CONTATO_PDPAS_PADRAO_ORDEM)]

    if is_necfm:
        linhas.append((
            "Informações técnicas, atestos, agendamentos e conclusão de serviços",
            f"NECFM - Engenharia Clínica — {contatos.get('NECFM', '')}"))
        return linhas

    if is_hrg and is_inv:
        linhas.append(("Aceite de Nota Fiscal / Gerente de Apoio Operacional HRG",
                       contatos.get("GERENTE_HRG", "")))
        linhas.append(("Recebimento de Produtos patrimoniais / agendamentos de "
                       "entrega/montagens", contatos.get("NPDOC_HRGINV", "")))
    elif is_hrg:
        linhas.append(("Aceite de Nota Fiscal / Recebimento de Produtos / agendamentos de "
                       "entrega/montagens", contatos.get("GERENTE_HRG", "")))
        linhas.append(("Farmácia (NFH)", contatos.get("NFH", "")))
        linhas.append(("Almoxarifado (NUAL)", contatos.get("NUAL_HRG", "")))
    elif is_inv:
        linhas.append(("Aceite de Nota Fiscal / Gerente de Apoio Operacional APS",
                       contatos.get("GERENTE", "")))
        linhas.append(("Recebimento de Produtos patrimoniais / agendamentos de "
                       "entrega/montagens", contatos.get("NPDOC_APSINV", "")))
    else:
        linhas.append(("Aceite de Nota Fiscal / Recebimento de Produtos / agendamentos de "
                       "entrega/montagens", contatos.get("GERENTE", "")))
        linhas.append(("Farmácia (NFL)", contatos.get("NFL", "")))
        linhas.append(("Almoxarifado (NUAL)", contatos.get("NUAL_APS", "")))

    return linhas


def montar_texto_email_ordem(tipo_planilha, origem, codigo_cotacao, ano, vencedores,
                              nome_signatario, matricula_signatario, cargo_signatario,
                              contatos=None):
    """Monta (assunto, corpo, destaques) do e-mail de ORDEM DE COMPRA,
    replicando os modelos reais da SES/PDPAS (APS, APS INV, HRG, HRG INV,
    e o caso especial da NECFM/Engenharia Clínica — SERVIÇO DE
    MANUTENÇÃO, sempre em vermelho, com contatos e itens 3/4 próprios).

    'vencedores' é uma lista de dicts {"nome_empresa", "itens_vencidos",
    "observacao"} — para o e-mail individual de um fornecedor, passe uma
    lista com um único vencedor; para o e-mail consolidado (TODOS OS
    FORNECEDORES), passe todos juntos: a tabela "Empresa Vencedora" ganha
    uma linha por fornecedor, e cada observação preenchida aparece
    identificada pelo nome da empresa correspondente.

    'destaques' devolvido é uma lista de (indice_linha_inicial,
    indice_linha_final, cor_hex, negrito) — usada só para colorir a
    prévia em tela; a cópia para a área de transferência é sempre texto
    puro (Tkinter não expõe formatação rica na área de transferência)."""
    contatos = contatos or CONTATOS_ORDEM_PADRAO
    tipo_norm = tipo_planilha.strip().upper()
    is_hrg = "HRG" in tipo_norm
    is_inv = "INV" in tipo_norm
    origem_norm = (origem or "").strip().upper()
    is_necfm = origem_norm == "NECFM"
    grupo = "HRG" if is_hrg else "APS"

    todos_itens_vencidos = [par for v in vencedores for par in v["itens_vencidos"]]
    objeto = gerar_objeto_automatico(
        [{"descricao": item.get("descricao") or ""} for item, _ in todos_itens_vencidos])

    assunto = f"ORDEM DE COMPRA COTAÇÃO Nº {codigo_cotacao}/{ano}-{tipo_norm} - {objeto}"

    if is_necfm:
        texto_cabecalho = ("COMUNICAÇÃO DE VENCEDOR e ORIENTAÇÕES: Processo de Aquisição por "
                           "Dispensa de Licitação SERVIÇO DE MANUTENÇÃO - PDPAS/SRSSU/SES-DF")
        cor_cabecalho = COR_CABECALHO_ORDEM_NECFM
    else:
        texto_cabecalho = f"{grupo} {'INVESTIMENTO' if is_inv else 'CUSTEIO'}"
        if is_hrg:
            cor_cabecalho = COR_CABECALHO_ORDEM_HRG_INV if is_inv else COR_CABECALHO_ORDEM_HRG
        else:
            cor_cabecalho = COR_CABECALHO_ORDEM_APS_INV if is_inv else COR_CABECALHO_ORDEM_APS

    linhas = []
    destaques = []

    linhas.append(texto_cabecalho)
    destaques.append((0, 0, cor_cabecalho, True))
    linhas.append("")

    linhas.append("Prezado(a) Fornecedor(a),")
    linhas.append("")
    linhas.append(
        "Temos o prazer de informar que sua empresa foi declarada vencedora do Pedido de "
        "Cotação referente ao processo de aquisição por dispensa de licitação, no âmbito do "
        "Programa de Descentralização Progressiva de Ações de Saúde – PDPAS/SES-DF.")
    linhas.append("")

    linhas.append("Empresa Vencedora")
    linhas.append("")
    linhas.append("Item(ns)\tEmpresa")
    for vencedor in vencedores:
        numeros_itens = ", ".join(
            str(item.get("numero_item") or "") for item, _ in vencedor["itens_vencidos"])
        linhas.append(f"{numeros_itens}\t{vencedor['nome_empresa']}")
    linhas.append("")

    observacoes_presentes = [(v["nome_empresa"], v.get("observacao") or "")
                             for v in vencedores if v.get("observacao")]
    if observacoes_presentes:
        indice_obs_inicio = len(linhas)
        if len(vencedores) == 1:
            # e-mail individual: mantém o texto puro, sem prefixar com o
            # nome da empresa (já dá pra saber de quem é, é o único vencedor)
            for trecho in observacoes_presentes[0][1].split("\n"):
                linhas.append(trecho)
        else:
            # e-mail consolidado (TODOS OS FORNECEDORES): cada observação
            # precisa dizer a qual empresa se refere
            linhas.append("Observações:")
            for nome_empresa, obs in observacoes_presentes:
                linhas.append(f"- {nome_empresa}: {obs}")
        destaques.append((indice_obs_inicio, len(linhas) - 1, COR_DESTAQUE_OBSERVACAO_ORDEM, True))
        linhas.append("")

    linhas.append("ORIENTAÇÕES PARA O FORNECIMENTO")
    linhas.append("")

    if is_necfm:
        linhas.append("1. Confirmação e Ordem de Fornecimento")
        linhas.append("")
        linhas.append(
            "    Solicitamos a confirmação imediata do recebimento deste e-mail e da respectiva "
            "Ordem de Fornecimento.")
        linhas.append("")
        linhas.append("2. Emissão da Nota Fiscal (NF)")
        linhas.append("")
        linhas.append("    A NF deve ser emitida no mesmo CNPJ constante da proposta comercial.")
        linhas.append(
            "    Detalhamento: A NF deve conter a descrição detalhada, o valor unitário e a "
            "quantidade de cada item, conforme esta Ordem de Compra.")
        linhas.append(
            "    Aprovação e Pagamento: Para evitar atrasos e cumprir as regras do PDPAS e da "
            "Lei nº 14.133/2021, a Nota Fiscal só deve ser emitida após a conclusão dos "
            "serviços e a devida aprovação/atesto pela nossa Engenharia Clínica. A NF só será "
            "atestada para pagamento após a entrega ou conclusão total.")
        linhas.append("    Retenções (Atenção!):")
        linhas.append(
            "    IR: Empresas não enquadradas no Simples Nacional terão o Imposto de Renda "
            "retido (1,2% para material e 4,8% para serviço), conforme IN RFB nº 2145/2023.")
        linhas.append(
            "    BRB (DF): Pagamentos para empresas do DF acima de R$ 5.000,00 serão efetuados "
            "somente em conta do BRB, conforme Decreto Distrital Nº 32.767/2011.")
        linhas.append("")
        linhas.append("3. Prazos e Local de Entrega/Serviço")
        linhas.append("")
        linhas.append(
            "    As entregas/serviços devem ser realizados rigorosamente dentro dos prazos "
            "estabelecidos na Ordem de Fornecimento, seguindo os princípios de legalidade e "
            "eficiência previstos na Lei nº 14.133/2021.")
        linhas.append("")
        linhas.append("4. Pagamento")
        linhas.append("")
        linhas.append(
            "    O pagamento será efetuado pela Região de Saúde Sul por TED/Transferência "
            "Eletrônica entre os dias 20 e 30 de cada mês, após o recebimento, conferência do "
            "material e o atesto do demandante (Engenharia Clínica).")
        linhas.append("")
        linhas.append("5. Encerramento do Exercício Financeiro")
        linhas.append("")
        linhas.append(TEXTO_ENCERRAMENTO_EXERCICIO_ORDEM)
        linhas.append("")
        linhas.append("6. Sanções")
        linhas.append("")
        linhas.append(TEXTO_SANCOES_ORDEM)
        linhas.append("")
    else:
        dados_unidade = _dados_institucionais_email_ordem(is_hrg)
        linhas.append("1. Confirmação e Ordem de Fornecimento")
        linhas.append("")
        linhas.append(
            "    Solicitamos a confirmação imediata do recebimento deste e-mail e da respectiva "
            "Ordem de Fornecimento, a fim de dar continuidade ao processo.")
        linhas.append("")
        linhas.append("2. Emissão da Nota Fiscal (NF)")
        linhas.append("")
        linhas.append(
            "    A Nota Fiscal deve ser emitida no mesmo CNPJ da proposta comercial e conter:")
        linhas.append(
            "    - Descrição detalhada do item, valor unitário e quantidade conforme esta "
            "Ordem de Compra;")
        linhas.append("    - Dados completos da unidade recebedora (abaixo);")
        linhas.append("    - Informações fiscais conforme legislação vigente.")
        linhas.append("")
        linhas.append("    Dados para emissão da NF e local de entrega de materiais:")
        linhas.append(f"    {dados_unidade['nome_unidade']}")
        linhas.append(f"    CNPJ: {dados_unidade['cnpj']}")
        linhas.append(f"    Inscrição Municipal: {dados_unidade['insc_municipal']}")
        linhas.append(f"    Endereço: {dados_unidade['endereco']}")
        if not is_hrg:
            linhas.append("    Ao lado da Igreja Nossa Senhora Aparecida")
            linhas.append(
                "    Coordenadas Google Maps: -16.007452594187583, -48.075027885099104")
            linhas.append("    Plus Code: XWVF+2X9 Gama, Brasília – DF")
        linhas.append(
            "    Horário de recebimento: segunda a quinta-feira, das 08h às 11h e das "
            "14h às 17h.")
        if is_inv:
            if not is_hrg:
                linhas.append(
                    "    Local de entrega (Farmácia e Almoxarifado): GAOAPS-SUL – endereço "
                    "acima.")
                linhas.append(
                    "    Entrega de bens patrimoniais: Núcleo de Patrimônio (NPDOC/APS) – "
                    "conforme agendamento.")
            else:
                linhas.append(
                    "    Entrega de bens patrimoniais: Núcleo de Patrimônio (NPDOC/HRG) – "
                    "conforme agendamento.")
        else:
            if not is_hrg:
                linhas.append(f"    Local de entrega: {dados_unidade['endereco']}")
            else:
                linhas.append(
                    "    Entrega de bens patrimoniais: Núcleo de Patrimônio (NPDOC/HRG) – "
                    "conforme agendamento.")
        linhas.append(
            "    Favor discriminar na Nota Fiscal o valor unitário e a quantidade de cada "
            "item, conforme a Ordem de Compra.")
        linhas.append("")
        linhas.append("3. Aprovação, Atesto e Pagamento")
        linhas.append("")
        linhas.append(
            "    O pagamento será realizado via TED ou Transferência Eletrônica, pela Região "
            "de Saúde Sul, entre os dias 20 e 30 de cada mês, após o recebimento e conferência "
            "do material/serviço e aprovação e atesto das áreas demandantes.")
        linhas.append("    Importante:")
        linhas.append("    - Não emitir boleto bancário.")
        linhas.append(
            "    - Os trâmites do processo de aquisição podem ser acompanhados pelo "
            "https://www.portalsei.df.gov.br/ em tempo real")
        linhas.append(
            "    - Empresas não enquadradas no Simples Nacional: retenção de IR (1,2% para "
            "material e 4,8% para serviço) conforme IN RFB nº 2145/2023.")
        linhas.append(
            "    - Empresas do DF: pagamentos acima de R$ 5.000,00 serão efetuados somente em "
            "conta do BRB, conforme Decreto Distrital nº 32.767/2011, art. 6º.")
        linhas.append("")
        linhas.append("4. Instalação e Montagem (quando aplicável)")
        linhas.append("")
        linhas.append(
            "    Caso o fornecimento envolva instalação e/ou montagem, o agendamento deve ser "
            "feito diretamente com o Chefe do Núcleo responsável (contato abaixo).")
        linhas.append("")
        linhas.append("5. Encerramento do Exercício Financeiro")
        linhas.append("")
        linhas.append(TEXTO_ENCERRAMENTO_EXERCICIO_ORDEM)
        linhas.append("")
        linhas.append("6. Sanções")
        linhas.append("")
        linhas.append(TEXTO_SANCOES_ORDEM)
        linhas.append("")

    linhas.append("CONTATOS ÚTEIS")
    linhas.append("")
    for assunto_contato, valor_contato in _contatos_uteis_email_ordem(
            is_hrg, is_inv, is_necfm, contatos):
        linhas.append(f"{assunto_contato}:\t{valor_contato}")
    linhas.append("")

    linhas.append("Agradecemos a parceria e colaboração.")
    linhas.append("Permanecemos à disposição para quaisquer esclarecimentos adicionais.")
    linhas.append("")
    linhas.append("Atenciosamente,")
    linhas.append("")

    linha_assinatura = nome_signatario or ""
    if matricula_signatario:
        linha_assinatura += f" - Matrícula {matricula_signatario}"
    linhas.append(linha_assinatura)
    if cargo_signatario:
        linhas.append(cargo_signatario)
    linhas.append("PDPAS - Programa de Descentralização Progressiva das Ações de Saúde")
    linhas.append("GEOF - Gerência de Orçamento e Finanças / DA / SRSSU")
    linhas.append("SUPERINTENDÊNCIA REGIONAL DE SAÚDE SUL CNPJ 00.394.700/0051-77")
    linhas.append("")
    linhas.append("Contato PREFERENCIALMENTE por e-mail")
    linhas.append("61 3449 7105")
    linhas.append("")
    linhas.append("HRG - Hospital Regional do Gama")
    linhas.append("Área Especial nº 01, Setor Central GAMA/DF. CEP. 72.405-901.")
    linhas.append("")
    linhas.append("APS SUL - Atenção Primária de Saúde Sul")
    linhas.append(
        "GAOAPS-SUL no Setor Oeste, SOE Q 12, Praça 1, Área Especial - Gama Oeste, "
        "CEP. 72.425-120.")

    corpo = "\n".join(linhas)
    return assunto, corpo, destaques


# ----------------------------------------------------------------------
# MIXIN — métodos de ação ligados aos botões do grupo E-mails
# ----------------------------------------------------------------------
class EmailMixin:
    """Métodos herdados por AppCotacao, ligados aos botões do grupo E-mails
    (E-mail Cotação, E-mail Ordem, FERIADOS)."""


    # ==================================================================
    # E-MAIL COTAÇÃO
    # ==================================================================
    def acao_email_cotacao(self):
        ano = self.entry_ano.get().strip()
        if not ano or len(ano) != 4:
            messagebox.showwarning("Atenção", "Informe o ANO com 4 dígitos.")
            return

        codigos = [e.get().strip() for e in self.entradas_cotacao if e.get().strip()]
        if not codigos:
            messagebox.showwarning(
                "Atenção", "Digite ao menos uma COTAÇÃO (código de 3 dígitos) para montar o e-mail.")
            return
        codigos_invalidos = [c for c in codigos if len(c) != 3]
        if codigos_invalidos:
            messagebox.showwarning(
                "Atenção",
                "Todas as cotações preenchidas precisam ter o código completo (3 dígitos).\n"
                f"Confira: {', '.join(codigos_invalidos)}")
            return

        caminho_arquivo = self._obter_caminho_arquivo()
        if not caminho_arquivo:
            messagebox.showwarning(
                "Atenção",
                "Não foi possível localizar o arquivo Matrix.\nClique em CONSULTAR MATRIX antes "
                "de montar o e-mail.")
            return

        try:
            linhas = buscar_cotacoes_na_planilha(caminho_arquivo, codigos, ano)
        except Exception as e:
            messagebox.showerror("Erro ao consultar", str(e))
            return

        grupos_cotacao = []
        codigos_sem_dados = []
        for codigo in codigos:
            codigo_padded = codigo.strip().zfill(3)
            linhas_do_codigo = [l for l in linhas if l["COTACAO"] == codigo_padded]
            if not linhas_do_codigo:
                codigos_sem_dados.append(codigo)
                continue
            sei = linhas_do_codigo[0].get("SEI") or ""
            objeto = gerar_objeto_automatico(
                [{"descricao": l.get("DESCRICAO") or ""} for l in linhas_do_codigo])
            grupos_cotacao.append({"codigo": codigo_padded, "sei": sei, "objeto": objeto})

        if codigos_sem_dados:
            messagebox.showwarning(
                "Atenção",
                "Nenhum item encontrado na aba 'Controle' para a(s) cotação(ões): "
                f"{', '.join(codigos_sem_dados)}.\nElas não entrarão no e-mail.")

        if not grupos_cotacao:
            messagebox.showwarning(
                "Atenção", "Nenhuma cotação com itens encontrados para montar o e-mail.")
            return

        dados = self._perguntar_dados_email_cotacao()
        if dados is None:
            return

        feriados = carregar_feriados()
        datas_feriados = obter_datas_feriados(feriados)
        data_recebimento = somar_dias_uteis(
            datetime.date.today(), dados["prazo_dias"], datas_feriados)
        data_recebimento_texto = data_recebimento.strftime("%d/%m/%Y")

        assunto, corpo, indice_linha_prazo = self._montar_texto_email_cotacao(
            codigos=[g["codigo"] for g in grupos_cotacao], ano=ano, grupos_cotacao=grupos_cotacao,
            publicacao=dados["publicacao"], horas=dados["horas"],
            data_recebimento_texto=data_recebimento_texto, exclusividade=dados["exclusividade"])

        self._mostrar_email_gerado(assunto, corpo, indice_linha_prazo, titulo="E-mail Cotação")

    def _perguntar_dados_email_cotacao(self):
        """Pergunta PUBLICAÇÃO, PRAZO (dias úteis, padrão 3), HORAS de
        recebimento e se é cotação com EXCLUSIVIDADE. Devolve um dict com
        as respostas, ou None se a janela for fechada sem confirmar."""
        janela = ctk.CTkToplevel(self)
        janela.title("Dados do E-mail de Cotação")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("520x460")
        janela.grab_set()

        ctk.CTkLabel(janela, text="PUBLICAÇÃO:", font=FONTE_SECAO, text_color=COR_TEXTO
                     ).pack(anchor="w", padx=16, pady=(16, 2))
        var_publicacao = tk.StringVar(value=PUBLICACOES_DISPONIVEIS[0])
        ctk.CTkOptionMenu(janela, values=PUBLICACOES_DISPONIVEIS, variable=var_publicacao,
                          font=FONTE_ENTRY, dropdown_font=FONTE_ENTRY, height=32
                          ).pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(janela, text="PRAZO para recebimento (dias úteis):", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(0, 2))
        entry_prazo = ctk.CTkEntry(janela, font=FONTE_ENTRY, height=32)
        entry_prazo.insert(0, "3")
        entry_prazo.pack(fill="x", padx=16, pady=(0, 10))
        entry_prazo.bind("<KeyRelease>", lambda e: self._aplicar_mascara_generica(
            entry_prazo, lambda t: formatar_apenas_digitos(t, 2)))

        ctk.CTkLabel(janela, text="HORAS de recebimento (HH:MM):", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(0, 2))
        entry_horas = ctk.CTkEntry(janela, font=FONTE_ENTRY, height=32, placeholder_text="18:00")
        entry_horas.pack(fill="x", padx=16, pady=(0, 10))
        entry_horas.bind("<KeyRelease>", lambda e: self._aplicar_mascara_generica(
            entry_horas, formatar_hora))

        var_exclusividade = tk.BooleanVar(value=False)
        ctk.CTkCheckBox(
            janela, text="Cotação com EXCLUSIVIDADE (inclui texto sobre Carta de Exclusividade)",
            variable=var_exclusividade, font=FONTE_TABELA, text_color=COR_TEXTO
        ).pack(anchor="w", padx=16, pady=(6, 10))

        resultado = {"valor": None}

        def confirmar():
            prazo_texto = entry_prazo.get().strip()
            horas_texto = entry_horas.get().strip()
            try:
                prazo_valor = int(prazo_texto)
            except ValueError:
                prazo_valor = 0
            if prazo_valor <= 0:
                messagebox.showwarning("Atenção", "Informe um PRAZO válido (em dias úteis).")
                return
            if len(horas_texto) != 5:
                messagebox.showwarning("Atenção", "Informe as HORAS completas (HH:MM).")
                return
            resultado["valor"] = {
                "publicacao": var_publicacao.get(),
                "prazo_dias": prazo_valor,
                "horas": horas_texto,
                "exclusividade": var_exclusividade.get(),
            }
            janela.destroy()

        ctk.CTkButton(janela, text="Gerar e-mail", font=FONTE_BOTAO, height=36,
                      fg_color=COR_DESTAQUE, command=confirmar).pack(pady=(10, 16))

        janela.protocol("WM_DELETE_WINDOW", janela.destroy)
        self.wait_window(janela)
        return resultado["valor"]

    def _montar_texto_email_cotacao(self, codigos, ano, grupos_cotacao, publicacao, horas,
                                     data_recebimento_texto, exclusividade):
        """Monta (ASSUNTO, CORPO, indice_linha_prazo). 'indice_linha_prazo'
        é a posição (0-based) da linha do item 2 dentro de CORPO, usada só
        para destacar essa linha visualmente na tela de prévia."""
        tipo = getattr(self, "_sufixo_planilha_atual", "")
        codigos_texto = ", ".join(codigos)
        prefixo = "PEDIDOS" if len(codigos) > 1 else "PEDIDO"
        assunto = (f"{prefixo} DE COTAÇÃO DE PREÇOS Nº {codigos_texto}/{ano} - "
                   f"PDPAS/{tipo} - {publicacao}.")

        linhas = [
            "Prezado(a) fornecedor(a),",
            "",
            "1 - Encaminhamos em anexo o Pedido de Cotação PDPAS/HRG e/ou PDPAS/APS citado no "
            "assunto do e-mail, para DISPENSA DE LICITAÇÃO em razão de valor, que será levada "
            "a efeito com fundamento no Inciso II, Art. 75 da Lei nº 14.133/21. Todos os "
            "documentos do processo poderão ser acompanhados pelo site "
            "https://www.portalsei.df.gov.br/",
            "",
        ]
        for grupo in grupos_cotacao:
            linhas.append(
                f"Cotação Nº {grupo['codigo']}/{ano}:\t{grupo['objeto']}\tSEI {grupo['sei']}")
        linhas.append("")

        indice_linha_prazo = len(linhas)
        linhas.append(
            f"2 - Prazo para recebimento das propostas: até às {horas} do dia "
            f"{data_recebimento_texto}")
        linhas.append("")

        linhas.extend(TEXTO_CORPO_BASE)

        if exclusividade:
            linhas.append("")
            linhas.append("")
            linhas.extend(TEXTO_EXCLUSIVIDADE)

        corpo = "\n".join(linhas)
        return assunto, corpo, indice_linha_prazo

    def _mostrar_email_gerado(self, assunto, corpo, indice_linha_destaque=None, titulo="E-mail"):
        janela = ctk.CTkToplevel(self)
        janela.title(titulo)
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("900x700")
        janela.grab_set()

        ctk.CTkLabel(janela, text="ASSUNTO:", font=FONTE_SECAO, text_color=COR_TEXTO
                     ).pack(anchor="w", padx=16, pady=(16, 2))
        linha_assunto = ctk.CTkFrame(janela, fg_color="transparent")
        linha_assunto.pack(fill="x", padx=16, pady=(0, 10))
        linha_assunto.grid_columnconfigure(0, weight=1)
        entry_assunto = ctk.CTkEntry(linha_assunto, font=FONTE_ENTRY, height=34)
        entry_assunto.grid(row=0, column=0, sticky="ew")
        entry_assunto.insert(0, assunto)
        entry_assunto.configure(state="readonly")
        ctk.CTkButton(linha_assunto, text="Copiar", width=80, height=34, font=FONTE_BOTAO,
                      command=lambda: self._copiar_para_area_transferencia(assunto)
                      ).grid(row=0, column=1, padx=(6, 0))

        ctk.CTkLabel(janela, text="CORPO DO E-MAIL:", font=FONTE_SECAO, text_color=COR_TEXTO
                     ).pack(anchor="w", padx=16, pady=(0, 2))
        caixa = ctk.CTkTextbox(janela, font=FONTE_ENTRY, wrap="word", height=480)
        caixa.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        caixa.insert("1.0", corpo)

        # Destaque visual (fonte maior, negrito, vermelho) na linha do item
        # 2 — só na prévia em tela, reproduzindo o estilo dos modelos reais
        # de e-mail enviados por PDPAS. É só um bônus visual: se a versão
        # do customtkinter não expuser o widget interno, o destaque é
        # simplesmente ignorado e o texto continua legível normalmente.
        if indice_linha_destaque is not None:
            try:
                texto_interno = caixa._textbox
                linha_tk = indice_linha_destaque + 1  # Tkinter Text conta linhas a partir de 1
                texto_interno.tag_add("destaque_prazo", f"{linha_tk}.0", f"{linha_tk}.end")
                texto_interno.tag_config("destaque_prazo", foreground="#ff4d4d",
                                         font=("Segoe UI", 15, "bold"))
            except Exception:
                pass

        caixa.configure(state="disabled")

        ctk.CTkButton(janela, text="Copiar corpo do e-mail", font=FONTE_BOTAO, height=34,
                      fg_color=COR_DESTAQUE,
                      command=lambda: self._copiar_para_area_transferencia(corpo)
                      ).pack(pady=(0, 8))
        ctk.CTkButton(janela, text="Fechar", font=FONTE_BOTAO, height=34,
                      command=janela.destroy).pack(pady=(0, 16))

    # ==================================================================
    # E-MAIL ORDEM
    # ==================================================================
    def acao_email_ordem(self):
        dados = self._montar_dados_ordens_ativas()
        if dados is None:
            return

        dados_gerais = self._perguntar_dados_gerais_email_ordem()
        if dados_gerais is None:
            return

        observacoes_por_fornecedor = self._perguntar_observacoes_email_ordem(
            dados["ordens_por_fornecedor"])
        if observacoes_por_fornecedor is None:
            return

        contatos = self.config_dados.get("contatos_ordem", CONTATOS_ORDEM_PADRAO)
        emails_gerados = []
        vencedores_todos = []
        for chave, (texto_vencedor, itens_vencidos) in dados["ordens_por_fornecedor"].items():
            vencedor = {
                "nome_empresa": texto_vencedor, "itens_vencidos": itens_vencidos,
                "observacao": observacoes_por_fornecedor.get(chave, ""),
            }
            vencedores_todos.append(vencedor)
            assunto, corpo, destaques = montar_texto_email_ordem(
                tipo_planilha=self._sufixo_planilha_atual, origem=dados["origem"],
                codigo_cotacao=dados["codigo_cotacao"], ano=dados["ano"],
                vencedores=[vencedor], nome_signatario=dados_gerais["nome"],
                matricula_signatario=dados_gerais["matricula"],
                cargo_signatario=dados_gerais["cargo"], contatos=contatos)
            emails_gerados.append({
                "fornecedor": texto_vencedor, "assunto": assunto, "corpo": corpo,
                "destaques": destaques,
            })

        # Além dos e-mails individuais (um por fornecedor), gera mais um
        # ao final reunindo TODOS os fornecedores vencedores no mesmo
        # e-mail (com as observações de cada um, quando existirem) — só
        # faz sentido quando há mais de um fornecedor vencedor.
        if len(vencedores_todos) > 1:
            assunto_todos, corpo_todos, destaques_todos = montar_texto_email_ordem(
                tipo_planilha=self._sufixo_planilha_atual, origem=dados["origem"],
                codigo_cotacao=dados["codigo_cotacao"], ano=dados["ano"],
                vencedores=vencedores_todos, nome_signatario=dados_gerais["nome"],
                matricula_signatario=dados_gerais["matricula"],
                cargo_signatario=dados_gerais["cargo"], contatos=contatos)
            emails_gerados.append({
                "fornecedor": "TODOS OS FORNECEDORES", "assunto": assunto_todos,
                "corpo": corpo_todos, "destaques": destaques_todos,
            })

        self._mostrar_emails_ordem_gerados(emails_gerados)

    def _perguntar_dados_gerais_email_ordem(self):
        """Pergunta quem assina os e-mails de Ordem (nome, matrícula e
        cargo opcionais) — campo livre porque vários usuários diferentes
        podem assinar, não é fixo. Vem pré-preenchido com o que foi salvo
        da última vez (self.config_dados["email_ordem_signatario"]) — ou,
        na primeira vez, com o OPERADOR selecionado na tela principal.
        Ao confirmar, o que estiver nos campos (alterado ou não) vira o
        novo padrão salvo em config_padrao.json, pronto para a próxima
        vez que o E-mail Ordem for usado."""
        salvo = self.config_dados.get("email_ordem_signatario", {})

        janela = ctk.CTkToplevel(self)
        janela.title("Dados do E-mail de Ordem")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("520x400")
        janela.grab_set()

        ctk.CTkLabel(janela, text="Nome de quem assina o e-mail:", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(16, 2))
        entry_nome = ctk.CTkEntry(janela, font=FONTE_ENTRY, height=32)
        entry_nome.insert(0, salvo.get("nome") or self.var_operador.get())
        entry_nome.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(janela, text="Matrícula (opcional):", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(0, 2))
        entry_matricula = ctk.CTkEntry(janela, font=FONTE_ENTRY, height=32,
                                       placeholder_text="Ex.: 1.436.681-9")
        entry_matricula.insert(0, salvo.get("matricula", ""))
        entry_matricula.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(janela, text="Cargo/Função (opcional):", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(0, 2))
        entry_cargo = ctk.CTkEntry(janela, font=FONTE_ENTRY, height=32,
                                   placeholder_text="Ex.: Gerente de Orçamento e Finanças")
        entry_cargo.insert(0, salvo.get("cargo", ""))
        entry_cargo.pack(fill="x", padx=16, pady=(0, 10))

        resultado = {"valor": None}

        def confirmar():
            nome = entry_nome.get().strip()
            if not nome:
                messagebox.showwarning("Atenção", "Informe o nome de quem assina o e-mail.")
                return
            dados_signatario = {
                "nome": nome,
                "matricula": entry_matricula.get().strip(),
                "cargo": entry_cargo.get().strip(),
            }
            self.config_dados["email_ordem_signatario"] = dados_signatario
            salvar_config(self.config_dados)
            resultado["valor"] = dados_signatario
            janela.destroy()

        ctk.CTkButton(janela, text="Continuar", font=FONTE_BOTAO, height=36,
                      fg_color=COR_DESTAQUE, command=confirmar).pack(pady=(10, 16))
        janela.protocol("WM_DELETE_WINDOW", janela.destroy)
        self.wait_window(janela)
        return resultado["valor"]

    def _perguntar_observacoes_email_ordem(self, ordens_por_fornecedor):
        """Tela com um campo de OBSERVAÇÃO opcional por fornecedor vencedor
        (ex.: pedir aprovação de marca/modelo antes do faturamento, como
        no modelo real 604/2026-PDPAS/APS - DISPLAY). Deixar em branco
        significa que aquele e-mail não terá o bloco de observação."""
        janela = ctk.CTkToplevel(self)
        janela.title("Observações por fornecedor (opcional)")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("900x560")
        janela.grab_set()

        ctk.CTkLabel(
            janela,
            text="Se algum fornecedor precisar de uma observação específica no e-mail de "
                 "Ordem (ex.: pedir confirmação de marca/modelo antes do faturamento), "
                 "escreva abaixo. Deixe em branco para não incluir observação.",
            font=FONTE_SECAO, text_color=COR_TEXTO, wraplength=860, justify="left"
        ).pack(anchor="w", padx=16, pady=(16, 8))

        scroll = ctk.CTkScrollableFrame(janela, fg_color="#161616", height=400)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        entradas = {}
        for chave, (texto_vencedor, itens_vencidos) in ordens_por_fornecedor.items():
            numeros = ", ".join(str(item.get("numero_item") or "") for item, _ in itens_vencidos)
            frame_forn = ctk.CTkFrame(scroll, fg_color="#1f1f1f", corner_radius=8)
            frame_forn.pack(fill="x", pady=4, padx=2)
            ctk.CTkLabel(frame_forn, text=f"{texto_vencedor} (item(ns) {numeros})",
                         font=("Segoe UI", 12, "bold"), text_color=COR_TEXTO, anchor="w"
                         ).pack(anchor="w", padx=10, pady=(8, 4))
            entry_obs = ctk.CTkEntry(frame_forn, font=FONTE_TABELA, height=30,
                                     placeholder_text="Observação (opcional)")
            entry_obs.pack(fill="x", padx=10, pady=(0, 8))
            entradas[chave] = entry_obs

        resultado = {"valor": None}

        def confirmar():
            resultado["valor"] = {chave: entry.get().strip() for chave, entry in entradas.items()}
            janela.destroy()

        ctk.CTkButton(janela, text="Gerar e-mails de Ordem", font=FONTE_BOTAO, height=36,
                      fg_color=COR_DESTAQUE, command=confirmar).pack(pady=(0, 16))
        janela.protocol("WM_DELETE_WINDOW", confirmar)
        self.wait_window(janela)
        return resultado["valor"]

    def _mostrar_emails_ordem_gerados(self, emails_gerados):
        """Tela de prévia com navegação (Fornecedor N de M): mostra o
        ASSUNTO e o CORPO do e-mail de Ordem de cada fornecedor vencedor,
        com o cabeçalho e a observação destacados na cor certa (só na
        tela — a cópia é sempre texto puro), e botões de Copiar."""
        janela = ctk.CTkToplevel(self)
        janela.title("E-mail Ordem")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("900x740")
        janela.grab_set()

        estado = {"indice": 0}

        label_cabecalho = ctk.CTkLabel(janela, text="", font=FONTE_SECAO, text_color=COR_TEXTO)
        label_cabecalho.pack(anchor="w", padx=16, pady=(16, 4))

        ctk.CTkLabel(janela, text="ASSUNTO:", font=FONTE_SECAO, text_color=COR_TEXTO
                     ).pack(anchor="w", padx=16, pady=(4, 2))
        linha_assunto = ctk.CTkFrame(janela, fg_color="transparent")
        linha_assunto.pack(fill="x", padx=16, pady=(0, 10))
        linha_assunto.grid_columnconfigure(0, weight=1)
        entry_assunto = ctk.CTkEntry(linha_assunto, font=FONTE_ENTRY, height=34)
        entry_assunto.grid(row=0, column=0, sticky="ew")
        botao_copiar_assunto = ctk.CTkButton(linha_assunto, text="Copiar", width=80, height=34,
                                             font=FONTE_BOTAO)
        botao_copiar_assunto.grid(row=0, column=1, padx=(6, 0))

        ctk.CTkLabel(janela, text="CORPO DO E-MAIL:", font=FONTE_SECAO, text_color=COR_TEXTO
                     ).pack(anchor="w", padx=16, pady=(0, 2))
        caixa = ctk.CTkTextbox(janela, font=FONTE_ENTRY, wrap="word", height=440)
        caixa.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        def redesenhar():
            i = estado["indice"]
            dados_email = emails_gerados[i]
            label_cabecalho.configure(
                text=f"Fornecedor {i + 1} de {len(emails_gerados)} — {dados_email['fornecedor']}")

            entry_assunto.configure(state="normal")
            entry_assunto.delete(0, "end")
            entry_assunto.insert(0, dados_email["assunto"])
            entry_assunto.configure(state="readonly")
            botao_copiar_assunto.configure(
                command=lambda: self._copiar_para_area_transferencia(dados_email["assunto"]))

            caixa.configure(state="normal")
            caixa.delete("1.0", "end")
            caixa.insert("1.0", dados_email["corpo"])
            try:
                texto_interno = caixa._textbox
                for indice, (inicio, fim, cor, negrito) in enumerate(dados_email["destaques"]):
                    tag = f"destaque_{indice}"
                    texto_interno.tag_add(tag, f"{inicio + 1}.0", f"{fim + 1}.end")
                    fonte = ("Segoe UI", 14, "bold") if negrito else ("Segoe UI", 14)
                    texto_interno.tag_config(tag, foreground=cor, font=fonte)
            except Exception:
                pass
            caixa.configure(state="disabled")

        def anterior():
            if estado["indice"] > 0:
                estado["indice"] -= 1
                redesenhar()

        def proximo():
            if estado["indice"] < len(emails_gerados) - 1:
                estado["indice"] += 1
                redesenhar()

        rodape = ctk.CTkFrame(janela, fg_color="transparent")
        rodape.pack(pady=(0, 8))
        ctk.CTkButton(rodape, text="◀ Fornecedor anterior", font=FONTE_BOTAO, height=34,
                      command=anterior).pack(side="left", padx=4)
        ctk.CTkButton(rodape, text="Copiar corpo do e-mail", font=FONTE_BOTAO, height=34,
                      fg_color=COR_DESTAQUE,
                      command=lambda: self._copiar_para_area_transferencia(
                          emails_gerados[estado["indice"]]["corpo"])
                      ).pack(side="left", padx=12)
        ctk.CTkButton(rodape, text="Fornecedor seguinte ▶", font=FONTE_BOTAO, height=34,
                      command=proximo).pack(side="left", padx=4)

        ctk.CTkButton(janela, text="Fechar", font=FONTE_BOTAO, height=34,
                      command=janela.destroy).pack(pady=(0, 16))

        redesenhar()

    # ==================================================================
    # FERIADOS — cadastro usado no cálculo do prazo em dias úteis
    # ==================================================================
    def acao_gerenciar_feriados(self):
        self._abrir_janela_feriados()

    def _abrir_janela_feriados(self):
        janela = ctk.CTkToplevel(self)
        janela.title("Feriados cadastrados")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("580x580")
        janela.grab_set()

        ctk.CTkLabel(
            janela,
            text="Cadastro de feriados — usados para calcular o prazo de recebimento das "
                 "propostas em dias úteis no E-mail Cotação (sábados, domingos e as datas "
                 "abaixo não contam como dia útil).",
            font=FONTE_SECAO, text_color=COR_TEXTO, wraplength=540, justify="left"
        ).pack(anchor="w", padx=16, pady=(16, 8))

        linha_form = ctk.CTkFrame(janela, fg_color="transparent")
        linha_form.pack(fill="x", padx=16, pady=(0, 10))

        ctk.CTkLabel(linha_form, text="Data:", font=FONTE_TABELA, text_color=COR_TEXTO
                     ).grid(row=0, column=0, padx=(0, 4))
        entry_data = ctk.CTkEntry(linha_form, font=FONTE_ENTRY, width=110, height=32,
                                  placeholder_text="DD/MM/AAAA")
        entry_data.grid(row=0, column=1, padx=(0, 10))
        entry_data.bind("<KeyRelease>", lambda e: self._aplicar_mascara_generica(
            entry_data, formatar_data))

        ctk.CTkLabel(linha_form, text="Feriado:", font=FONTE_TABELA, text_color=COR_TEXTO
                     ).grid(row=0, column=2, padx=(0, 4))
        entry_texto = ctk.CTkEntry(linha_form, font=FONTE_ENTRY, width=200, height=32,
                                   placeholder_text="Ex.: Independência")
        entry_texto.grid(row=0, column=3, padx=(0, 10))

        scroll = ctk.CTkScrollableFrame(janela, fg_color="#161616", height=380)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        def redesenhar_lista():
            for w in scroll.winfo_children():
                w.destroy()
            feriados = sorted(carregar_feriados(), key=self._chave_ordenacao_feriado)
            if not feriados:
                ctk.CTkLabel(scroll, text="Nenhum feriado cadastrado ainda.",
                             font=FONTE_TABELA, text_color=COR_TEXTO_SEC
                             ).pack(anchor="w", padx=8, pady=8)
                return
            for feriado in feriados:
                linha = ctk.CTkFrame(scroll, fg_color="#1f1f1f", corner_radius=6)
                linha.pack(fill="x", pady=3, padx=2)
                ctk.CTkLabel(linha, text=feriado.get("data", ""), font=FONTE_TABELA,
                             text_color=COR_TEXTO, width=90
                             ).pack(side="left", padx=(8, 6), pady=6)
                ctk.CTkLabel(linha, text=feriado.get("feriado", ""), font=FONTE_TABELA,
                             text_color=COR_TEXTO_SEC, anchor="w"
                             ).pack(side="left", fill="x", expand=True, pady=6)
                ctk.CTkButton(linha, text="Remover", width=80, height=26, font=FONTE_BOTAO,
                              fg_color="#7a2626",
                              command=lambda f=feriado: remover(f)).pack(side="right", padx=8)

        def adicionar():
            data_texto = entry_data.get().strip()
            texto = entry_texto.get().strip()
            if len(data_texto) != 10:
                messagebox.showwarning("Atenção", "Informe a DATA completa (DD/MM/AAAA).")
                return
            if not texto:
                messagebox.showwarning("Atenção", "Informe uma descrição curta para o feriado.")
                return
            feriados = carregar_feriados()
            feriados.append({"data": data_texto, "feriado": texto})
            salvar_feriados(feriados)
            entry_data.delete(0, "end")
            entry_texto.delete(0, "end")
            redesenhar_lista()

        def remover(feriado_alvo):
            feriados = carregar_feriados()
            feriados = [f for f in feriados
                       if not (f.get("data") == feriado_alvo.get("data")
                               and f.get("feriado") == feriado_alvo.get("feriado"))]
            salvar_feriados(feriados)
            redesenhar_lista()

        ctk.CTkButton(linha_form, text="Adicionar", font=FONTE_BOTAO, height=32,
                      fg_color=COR_DESTAQUE, command=adicionar).grid(row=0, column=4, padx=(4, 0))

        redesenhar_lista()

        ctk.CTkButton(janela, text="Fechar", font=FONTE_BOTAO, height=34,
                      command=janela.destroy).pack(pady=(0, 16))

    @staticmethod
    def _chave_ordenacao_feriado(feriado):
        texto = (feriado.get("data") or "").strip()
        partes = texto.split("/")
        if len(partes) == 3:
            try:
                d, m, a = partes
                return (int(a), int(m), int(d))
            except ValueError:
                pass
        return (9999, 99, 99)

    @staticmethod
    def _aplicar_mascara_generica(entry, func_mascara):
        texto = entry.get()
        formatado = func_mascara(texto)
        if formatado != texto:
            entry.delete(0, "end")
            entry.insert(0, formatado)
            entry.icursor("end")