# Relatório de continuidade — GEOF-CODEX / CONTROLE DE COTAÇÕES

**Data de referência:** 12/11/2026.

Versão em Markdown do [relatório em PDF](Relatorio_Continuidade_GEOF_CODEX.pdf). O conteúdo registra o marco técnico indicado no relatório; não substitui a consulta ao código e ao histórico mais recentes.

## GEOF-CODEX | CONTROLE DE COTAÇÕES

RELATÓRIO DE CONTINUIDADE PARA OUTRA IA
Data de referência: 12/11/2026

### 1. Finalidade e alcance

Este documento permite retomar a manutenção do aplicativo sem começar do zero. O usuário informa que houve outras conversas sobre o mesmo sistema, identificadas pelos termos GEOF-CODEX ou CONTROLE DE COTAÇÕES. Esses nomes devem ser tratados como referências ao mesmo projeto.

Fontes: conversa atual; README.md e AGENTS.md consultados no GitHub; histórico de commits da main consultado em 12/11/2026. As outras conversas não foram acessadas diretamente. Mudanças anteriores são descritas como evidências do histórico do repositório, sem atribuir a elas decisões não verificadas.

Repositório:
https://github.com/oberdan-alves/GEOF-CODEX

Branch de trabalho: main.
Último commit da main na consulta:
ba6eba9d8726143a8f43338ae5d72c73913a4db0

Os quatro commits desta conversa estão publicados. Nenhuma alteração de funcionalidade desta conversa ficou pendente de envio.

### 2. O que é o aplicativo

Aplicativo desktop em Python, com interface CustomTkinter/Tkinter e uso em Windows. Lê planilhas Matrix, organiza cotações e propostas de fornecedores, compara preços, registra decisões por item e gera pedidos, mapas, ordens e textos de e-mail.

Fluxo principal:

- Selecionar pasta GEOF, planilha, ano e cotação ativa.
- Consultar os itens da cotação.
- Criar pasta e gerar Pedido de Cotação.
- Identificar empresas e cadastrar propostas.
- Analisar pesquisas, preços, retiradas, desfechos e vencedores.
- Gerar Mapa Comparativo e Ordens de Compra.
- Preparar textos dos e-mails e administrar feriados.

A consulta remota ao GitHub foi realizada em 12/11/2026. O ambiente de arquivos foi restabelecido para gerar este PDF. Os resultados de testes citados são históricos; não houve nova execução dos testes do aplicativo para a emissão do documento.

## Arquitetura e dados compartilhados

### 3. Onde cada parte do código fica

GeradorCotacao.py
Entrada do programa, tela principal, cores, fontes, utilitários, leitura de itens, configurações, cabeçalho institucional e conversão PDF para JPG. A base _AppCotacaoBase é combinada com os mixins por criar_app().

Cotacao.py
CotacaoMixin: criação das pastas, Pedido de Cotação e ajustes do pedido.

Mapa.py
MapaMixin: leitura de fornecedores, listagem das empresas, cadastro e prévia das propostas, ranking e cálculos dos totais.

AdjudicacaoOrdem.py
AdjudicacaoOrdemMixin: cancelamentos, pesquisas e revisão dos preços, decisões e vencedores por item, Mapa Comparativo e Ordem de Compra.

Emails.py
EmailMixin: preparação dos textos de e-mail de cotação e ordem e cadastro de feriados. Não confundir preparar texto com enviar mensagem.

Outros arquivos importantes:

- caminhos.py: resolução de caminhos e armazenamento compartilhado.
- Mensagens.py: diálogos do aplicativo.
- criar_executavel.py: verificar o procedimento atual de empacotamento.
- assets/: brasão e imagens dos documentos.
- tests/: regressões automatizadas.

### 4. Dados e caminhos que precisam ser preservados

Cotações:
GEOF/<tipo>/PDPAS <ano>/COTACAO/<código>

Cadastro e decisões:
MAPA_EMPRESAS_<código>.json
Preservar fornecedores, decisoes_itens e dados_mapa ao editar propostas.

AJUSTES_PEDIDO_<código>.json
Cancelamentos e observações específicos do Pedido; não aplicar essas alterações automaticamente às decisões do mapa ou da ordem.

EMPRESAS_LISTADAS_<código>.json
Lista persistida usada também na indicação de conclusão das etapas.

Segundo o README atualizado, configurações, operadores, contatos, feriados e recursos usam uma pasta irmã de GEOF:
Planilhas Geradoras Cotação Pagamento Prestação/GeradorCotacao

Na primeira utilização, recursos locais podem ser copiados sem substituir os compartilhados existentes. Sem acesso à pasta compartilhada, a gravação deve informar erro, sem usar uma cópia local silenciosa.

O README ainda contém trechos antigos sobre localização de config_padrao.json e funções pendentes. Conferir código e histórico antes de usar essas seções como especificação atual.

## Alterações desta conversa: valores e layout

### 5. Precisão dos valores - commit 4dcfed8

Problema: o unitário admitia quatro casas decimais, mas cada total era arredondado para duas casas ao salvar. A soma produzia diferenças em relação a quantidade x unitário, propagadas para adjudicação e ordem.

Requisito: TOTAL GERAL DA PROPOSTA deve ser a soma dos produtos quantidade x unitário, preservando as casas significativas dos itens.

Implementação em Mapa.py:

- calcular_total_proposta(quantidade, unitario) multiplica usando Decimal a partir de strings.
- formatar_total_proposta(valor) exibe ao menos duas casas e preserva casas significativas.
- Removido round(..., 2) do salvamento de cada total.
- Ranking e montagem dos dados da adjudicação recalculam produtos, inclusive para cadastros antigos cujo campo total estava arredondado.
- A prévia soma os produtos em Decimal.
- Ordem de Compra recalcula os produtos e usa soma decimal.
- Documentos deixam de limitar os unitários e totais às duas casas.

Exemplo de aceitação:
3 x 1,2345 = 3,7035
Três itens desse valor somam 11,1105.
Não salvar cada item como 3,70 antes de somar.

Limite técnico importante: há conversões do total para float para manter o JSON e estruturas existentes compatíveis. Não afirmar que todo o armazenamento financeiro foi migrado para Decimal. Ao ampliar a precisão ou a escala, revisar também serialização, releitura e formatação.

A interface recebeu fonte maior, campos sob os títulos e remoção de R$ nos totais da tela. Não remover indiscriminadamente R$ dos documentos: a apresentação monetária em PDFs foi preservada onde necessário.

### 6. Descrição e coluna UN - commit 81df01f

Pedido do usuário baseado em imagens da tela:

- DESCRIÇÃO deve utilizar o espaço entre CÓDIGO e UN.
- Na prévia, aproximar o texto dessas colunas.
- Ampliar UN para que Serviço não quebre em duas linhas.

Aplicado em formulário e prévia:

- Removido o corte da descrição em 55 caracteres.
- Texto completo, alinhado à esquerda.
- Coluna de descrição expande e quebra conforme a largura real.
- Ajuste em <Configure>, respeitando escala do CustomTkinter.
- UN ampliada de 50 para 85 na definição das colunas.

Manter essa adaptação ao redimensionar e ao maximizar. Testar também a linha ativa em negrito e descrições longas.

## Alterações desta conversa: dicas e teclado

### 7. Dicas dos botões - commit a631174

Na tela CONTROLE DE COTAÇÕES, exibir explicação da ação quando o mouse fica sobre um botão. Se a cor já indicar conclusão da etapa, acrescentar essa informação.

Implementação em GeradorCotacao.py:

- DICAS_BOTOES_CONTROLE: descrições dos botões.
- DicaBotao: janela de dica sem foco, exibida após 400 ms.
- Ocultação ao sair, clicar ou destruir o botão.
- _criar_botao_com_dica(): criação e ligação da dica.
- _texto_dica_botao(): descrição consultada a cada passagem.
- _atualizar_cores_etapas() atualiza _etapas_concluidas.

Texto adicional:
Ação já executada para a cotação ativa (indicada pela cor do botão).

A mesma evidência determina cor e dica, com base nos dados da cotação ativa. Não marcar como concluído apenas porque o botão foi clicado.

Etapas monitoradas por arquivos e cadastro:
CRIAR PASTA COTAÇÃO, PEDIDO, LISTAR EMPRESAS,
INCLUIR EMPRESAS MAPA, ADJUDICAÇÃO e ORDEM.

As dicas também cobrem consulta, pasta, contatos, operador, cópias de SIS/SEI/código e os demais botões dos grupos. A indicação de conclusão deve acompanhar mudança de cotação, ano ou tipo de planilha.

### 8. Seleção dos campos e Enter - commit ba6eba9

No Cadastro de Propostas por Fornecedor:

- Ao clicar no campo ou receber foco, selecionar todo o conteúdo.
- Permitir substituir o valor apenas digitando.
- Enter normal e do teclado numérico avançam na mesma ordem do Tab.
- Tab permanece com seu comportamento original.

Implementação em Mapa.py:
_configurar_navegacao_campo_proposta(entry, idx=None)

Aplicada à data e aos campos quantidade, unitário, marca e prazo.
A seleção é agendada com after_idle para ocorrer após o posicionamento do cursor pelo clique. Usa select_range(0, "end") e verifica se o campo ainda existe. O foco continua ativando o destaque da linha.

<Return> e <KP_Enter> usam event.widget.tk_focusNext().focus_set() e retornam "break".
Mantidos os callbacks de saída para normalizar quantidade/unitário e atualizar os totais.

Verificar manualmente cliques, Tab, Enter, teclado numérico, passagem entre linhas e campos com máscaras. Os testes automatizados não substituem verificação visual no Windows.

## Histórico anterior identificado no GitHub

### 9. Trabalhos anteriores: evidências do repositório

Esta seção resume assuntos encontrados nos títulos dos commits. Não é uma transcrição das outras conversas e não comprova todos os detalhes de suas regras de negócio. Consultar o diff do commit e o código atual antes de alterar.

Interface e propostas:

- 4aa1535: CNPJ junto ao fornecedor no cadastro.
- cab9f8f: data padrão por cotação e cadastro maximizado.
- 9a85b99: tela Controle de Cotações maximizada.
- f5363a5: colunas alinhadas, repetição de dados com desfazer e unitários com quatro casas.
- f8195d2 e 4634b3f: ampliação de fontes e destaque da linha em edição.
- 6f902fc: CNPJ na seleção de fornecedores.
- 945a117: cores para etapas concluídas.
- b4a4ba8 e a92bb54: limpeza e edição das caixas de cotação.
- 331b875: abertura da planilha selecionada em Consultar Matrix.

Pesquisas, revisão e persistência:

- a4e04d2: integração das pesquisas e revisão à adjudicação.
- 73160df: fluxo de pesquisa, revisão e vencedor após cancelamentos.
- 6b5e613: retirada de propostas com justificativa.
- ed444de: retiradas por limites e classificação dos preços.
- 530f9e4: retorno de fornecedores retirados na revisão.
- f9eebc3: simulador de referência e teto de negociação.
- 279de8e: persistência das telas e limpeza isolada.
- c9339c4: preservação das decisões ao atualizar propostas.
- 7824f91: desfechos por item integrados aos documentos.
- b1f4103: relatório explicativo de pesquisa, com revisão e texto persistido.
- 63cb104, eea394f, 9272c36 e c810b16: ajustes da metodologia, pesquisas incompletas, numeração e preservação do texto editado.

Documentos, infraestrutura e comunicações:

- 32261a9: diagramação da adjudicação e destaque somente do nome do vencedor.
- 96edb2f: publicação e rodapé do mapa e critério da referência.
- 96d0b16, c8c20f7, 4f7a4aa e 0310b92: modelos APS/HRG, contatos por setor e versões de investimento.
- ca28ab2: informações técnicas e aceite de NF na Ordem HRG.
- 7f8dcae: observações antes dos dados fiscais da ordem.
- a10651c: cláusulas e edição do texto do e-mail de cotação.
- 8003a52: configurações e recursos compartilhados derivados de GEOF.
- 3d615d4: normalização de caminhos de rede.

Existem commits de atualização pela versão local. Atribuir funcionalidades somente após conferir o conteúdo, pois o título desses commits é genérico.

## Como retomar e validar com segurança

### 10. Regras e procedimento de continuidade

Ler AGENTS.md vigente. O arquivo consultado determina:

- Trabalhar na main; não criar branch alternativa sem autorização.
- Confirmar branch e sincronizar antes de editar.
- Alterar os arquivos originais e preservar funcionalidades fora do escopo.
- Revisar diff e executar os testes disponíveis.
- As solicitações incluem autorização de commit e publicação quando tecnicamente permitido.
- Relatar arquivos, testes, commit e resultado efetivo do envio.
- Se houver bloqueio de autorização, explicar e aguardar; não contornar.

Nesta conversa, a revisão automática bloqueou alguns pushes apesar do AGENTS.md. O usuário autorizou explicitamente os commits a631174 e ba6eba9, e ambos foram enviados depois. Não confundir commit local com publicação.

Para iniciar uma nova alteração:
1) Conferir git status, branch, remoto e alterações locais.
2) Sincronizar a main sem descartar trabalho existente.
3) Ler somente os módulos relacionados e seus testes.
4) Identificar as regras já preservadas por alterações anteriores.
5) Fazer ajuste no código original, revisar e testar.
6) Confirmar o resultado do push, se houver publicação autorizada.
7) Informar se o executável também foi recompilado.

Comandos de referência:
pip install -r requirements.txt
python GeradorCotacao.py
python -m unittest discover -s tests
git diff --check

Dependências conhecidas:
customtkinter, openpyxl, pandas, reportlab, pillow, pymupdf.
PDF para JPG usa PyMuPDF/fitz, sem exigir Poppler.

### 11. Testes e estado de entrega

4dcfed8: precisão e layout; 84 testes na validação correspondente.
81df01f: descrição e UN; 84 testes.
a631174: dicas e conclusão; 86 testes.
ba6eba9: seleção e Enter; 87 testes.

Arquivos principais alterados nesta conversa:
Mapa.py; AdjudicacaoOrdem.py; GeradorCotacao.py.
Testes alterados:
tests/test_cadastro_propostas.py
tests/test_cores_etapas.py
tests/test_botao_adjudicacao.py

A última execução registrada, antes de publicar ba6eba9, passou nos 87 testes. Essa é evidência histórica; não foi uma nova execução na emissão deste relatório.

Os testes de tela usam objetos simulados em vários casos. Validar visualmente no Windows: escala de exibição, descrições longas, foco, seleção, dicas e navegação.

Alterar .py e publicar no GitHub não atualiza um .exe existente. Nesta conversa não foi feita nova compilação do executável Windows. Verificar criar_executavel.py e produzir uma versão nova quando necessário.

## Instrução pronta para a próxima IA

### 12. Contexto para entregar junto com os arquivos

Você está dando continuidade ao aplicativo desktop Python CONTROLE DE COTAÇÕES, do repositório oberdan-alves/GEOF-CODEX. Referências anteriores a GEOF-CODEX e CONTROLE DE COTAÇÕES dizem respeito ao mesmo projeto.

O trabalho foi desenvolvido em várias conversas. Este relatório reúne a conversa disponível e evidências do GitHub, mas não substitui todos os diálogos anteriores. Consulte a main, os testes, README.md, AGENTS.md e o histórico antes de concluir que uma função está ausente ou deve ser refeita.

Último commit confirmado em 12/11/2026:
ba6eba9d8726143a8f43338ae5d72c73913a4db0

Preserve especialmente:

- Soma dos produtos quantidade x unitário sem arredondar cada item para centavos.
- Unitários com até quatro casas no cadastro.
- Compatibilidade dos cadastros e decisões já salvos.
- Descrições completas, com largura adaptável entre CÓDIGO e UN.
- Coluna UN ampliada no formulário e na prévia.
- Dicas dos botões e aviso de conclusão conforme a cotação ativa.
- Seleção integral dos campos e Enter equivalente ao Tab.
- Pesquisas, retiradas, justificativas, desfechos e vencedores por item.
- Dados compartilhados derivados da pasta GEOF.
- Modelos institucionais APS/HRG e versões de investimento.
- Textos editados e decisões persistidas sem sobrescrita indevida.

Não trate os placeholders antigos do README como diagnóstico atual. Por exemplo, ACESSAR PROPOSTAS e os e-mails já têm implementação identificada na conversa/código anterior; o botão SALVAR do 3º Passo ainda tinha uma ação adicional marcada como stub, devendo ser reavaliado antes de assumir mudança de comportamento.

Não reescreva o aplicativo inteiro para um ajuste pontual. Confirme as instruções da main, trabalhe nos arquivos originais e informe testes, commit, publicação e compilação separadamente.

### 13. Referências verificáveis

Repositório e histórico:
https://github.com/oberdan-alves/GEOF-CODEX
https://github.com/oberdan-alves/GEOF-CODEX/commits/main

Instruções e documentação:
https://github.com/oberdan-alves/GEOF-CODEX/blob/main/AGENTS.md
https://github.com/oberdan-alves/GEOF-CODEX/blob/main/README.md

Commits detalhados desta conversa:
4dcfed8 - precisão e alinhamento
81df01f - descrição e UN
a631174 - dicas dos botões
ba6eba9 - seleção e Enter

Ao retomar, compare a main com esse marco. Caso existam commits posteriores, o estado mais recente deve ser revisado antes de aplicar este relatório como orientação técnica.
