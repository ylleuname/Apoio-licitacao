import cloudscraper
import pandas as pd
import os
import time
import requests
import re
import json
from datetime import datetime
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


def obter_valor_multiplas_chaves(dicionarios, chaves):
    """Vasculha múltiplos dicionários buscando por chaves em qualquer variação de escrita."""
    chaves_lower = [c.lower() for c in chaves]
    for d in dicionarios:
        if not isinstance(d, dict):
            continue
        for k, v in d.items():
            if k.lower() in chaves_lower:
                if v is not None and str(v).strip().lower() not in ('', 'none', 'null'):
                    return v
    return ""


def formatar_data(data_texto):
    """Converte variações de datas da API para o formato DD/MM/YYYY HH:MM."""
    if not data_texto or str(data_texto).strip().lower() in ("", "none", "null"):
        return "Não Informada"
    try:
        dt = pd.to_datetime(data_texto).tz_localize(None)
        return dt.strftime('%d/%m/%Y %H:%M')
    except:
        return str(data_texto)


def baixar_arquivos_pncp(cnpj: str, ano: str, seq: str, orgao: str, scraper, pasta_raiz: str):
    orgao_limpo = re.sub(r'[\\/*?:"<>|]', "", str(orgao)).strip()[:60]
    if not orgao_limpo:
        orgao_limpo = f"Orgao_{cnpj}"

    nome_pasta = f"{orgao_limpo}_Seq_{seq}"
    pasta_destino = os.path.join(pasta_raiz, nome_pasta)
    url_lista_arquivos = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}/arquivos"

    try:
        res = scraper.get(url_lista_arquivos, headers={'Accept': 'application/json'}, timeout=15, allow_redirects=True)
        if res.status_code != 200:
            return

        arquivos = res.json()
        if not arquivos:
            return

        os.makedirs(pasta_destino, exist_ok=True)

        for arq in arquivos:
            titulo = arq.get('tituloDocumento') or arq.get('titulo') or "Documento"
            url_down = arq.get('url') or arq.get('linkArquivo') or arq.get('urlDocumento')

            if not url_down:
                seq_arq = arq.get('sequencialDocumento') or arq.get('sequencialArquivo') or arq.get('sequencial')
                if seq_arq:
                    url_down = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}/arquivos/{seq_arq}"

            if not url_down:
                continue

            titulo_limpo = re.sub(r'[\\/*?:"<>|]', "", str(titulo)).strip()[:80]
            if not titulo_limpo:
                titulo_limpo = f"Documento_{seq_arq}"

            if not '.' in titulo_limpo[-5:]:
                titulo_limpo += ".pdf"

            caminho_arquivo = os.path.join(pasta_destino, titulo_limpo)
            if os.path.exists(caminho_arquivo):
                continue

            try:
                doc_res = scraper.get(url_down, stream=True, allow_redirects=True, timeout=60)
                if doc_res.status_code == 200:
                    content_type = doc_res.headers.get('Content-Type', '')
                    if 'application/json' in content_type:
                        continue
                    with open(caminho_arquivo, 'wb') as f:
                        for chunk in doc_res.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
            except Exception:
                pass
    except Exception:
        pass


def buscar_editais_pncp(palavras_chave: list, pasta_raiz: str, app_gui=None, baixar_arquivos=True) -> pd.DataFrame:
    base_url = "https://pncp.gov.br/api/search"
    resultados_totais = []

    arquivo_excel = os.path.join(pasta_raiz, "editais_monitorados.xlsx")
    links_ja_processados = set()

    if os.path.exists(arquivo_excel):
        try:
            df_antigo = pd.read_excel(arquivo_excel)
            if 'link' in df_antigo.columns:
                links_ja_processados = set(df_antigo['link'].astype(str).tolist())
        except Exception:
            pass

    scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows', 'desktop': True})
    headers_detalhe = {'Accept': 'application/json'}
    status_abertos = ['recebendo_proposta']

    for keyword in palavras_chave:
        kw_clean = keyword.lower().strip()

        # Define variações diretas para evitar perdas por grafia ou hífen
        variacoes_kw = [kw_clean]
        if kw_clean == "videowall":
            variacoes_kw.extend(["video wall", "video-wall"])
        elif kw_clean == "cftv":
            variacoes_kw.extend(["circuito fechado de tv", "circuito fechado de televisão"])
        elif kw_clean == "fibra óptica":
            variacoes_kw.extend(["fibra optica", "fibra-optica"])

        for status_slug in status_abertos:
            pagina_atual = 1

            while True:
                if app_gui:
                    texto_status = f"A procurar: '{keyword}' (Página {pagina_atual})..." if pagina_atual > 1 else f"A procurar: '{keyword}'..."
                    app_gui.atualizar_status(texto_status)

                params = {
                    'q': keyword,
                    'tipos_documento': 'edital',
                    'status': status_slug,
                    'ordenacao': '-data',
                    'pagina': pagina_atual,
                    'tam_pagina': 50
                }

                max_tentativas = 3
                sucesso_requisicao = False
                dados = []

                for tentativa in range(1, max_tentativas + 1):
                    try:
                        response = scraper.get(base_url, params=params, timeout=60)
                        response.raise_for_status()
                        dados = response.json().get('items', [])
                        sucesso_requisicao = True
                        break
                    except Exception as e:
                        print(f"[ERRO SERVIDOR SEARCH] Tentativa {tentativa}: {e}")
                        if tentativa < max_tentativas:
                            time.sleep(5)

                if not sucesso_requisicao or not dados:
                    break

                for licitacao in dados:
                    cnpj = licitacao.get('orgao_cnpj')
                    ano = licitacao.get('ano')
                    seq = licitacao.get('numero_sequencial')

                    link_acesso = f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq}" if (
                                cnpj and ano and seq) else licitacao.get('itemUrl', 'Link indisponível')

                    if link_acesso in links_ja_processados:
                        continue

                    title = str(licitacao.get('title') or "").strip()
                    desc = str(licitacao.get('description') or "").strip()
                    objeto_texto = f"{title} {desc}".strip()
                    objeto_lower = objeto_texto.lower()

                    # --- FILTRO RIGOROSO DE PALAVRA-CHAVE (POST-FILTERING) ---
                    # Elimina falsos positivos do PNCP que não contêm o termo exato no resumo do objeto
                    if not any(var in objeto_lower for var in variacoes_kw):
                        continue

                    orgao = licitacao.get('orgao_nome', 'Não Informado')

                    # --- REQUISIÇÃO À API DE DETALHES ---
                    json_detalhe = {}
                    if cnpj and ano and seq:
                        detalhe_url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}"
                        for tentativa_detalhe in range(3):
                            try:
                                time.sleep(0.2)
                                res_detalhe = scraper.get(detalhe_url, headers=headers_detalhe, timeout=10,
                                                          allow_redirects=True)
                                if res_detalhe.status_code == 200:
                                    json_detalhe = res_detalhe.json()
                                    break
                                elif res_detalhe.status_code == 429:
                                    time.sleep(2)
                            except Exception as e:
                                print(f"[ERRO DETALHES] {e}")
                                time.sleep(1)

                    unidade_orgao = json_detalhe.get('unidadeOrgao') or {}

                    # --- 1. FILTRO ESFERA (Ignorar Municipal) ---
                    esfera_id = obter_valor_multiplas_chaves([json_detalhe.get('orgaoEntidade', {}), licitacao],
                                                             ['esferaId', 'esfera_id'])
                    if not esfera_id:
                        if 'MUNICIPAL' in orgao.upper() or 'PREFEITURA' in orgao.upper():
                            esfera_id = 'M'
                    if str(esfera_id).upper() == 'M':
                        continue

                        # --- 2. MODALIDADE E MODO DE DISPUTA ---
                    mod_bruta = str(obter_valor_multiplas_chaves(
                        [json_detalhe, licitacao],
                        ['modalidade_licitacao_nome', 'modalidadeNome', 'modalidade_nome', 'modalidade',
                         'modalidade_contratacao_nome']
                    )).strip()

                    disp_bruta = str(obter_valor_multiplas_chaves(
                        [json_detalhe, licitacao],
                        ['modo_disputa_nome', 'modoDisputaNome', 'modo_disputa', 'mododisputanome', 'tipo_disputa_nome']
                    )).strip()

                    texto_comb_lower = f"{mod_bruta} {disp_bruta} {title}".lower()

                    # BLOQUEIO DE DISPENSA COM DISPUTA
                    if "dispensa" in texto_comb_lower and "disputa" in texto_comb_lower:
                        continue

                    if mod_bruta:
                        modalidade_formatada = mod_bruta.title()
                        if disp_bruta and disp_bruta.lower() not in mod_bruta.lower():
                            modalidade_formatada += f" - {disp_bruta.title()}"
                    elif disp_bruta:
                        modalidade_formatada = disp_bruta.title()
                    else:
                        title_upper = title.upper()
                        if "PREGÃO" in title_upper:
                            modalidade_formatada = "Pregão Eletrônico" if "ELETRÔNICO" in title_upper else "Pregão"
                        elif "CONCORRÊNCIA" in title_upper:
                            modalidade_formatada = "Concorrência"
                        elif "DISPENSA" in title_upper:
                            modalidade_formatada = "Dispensa"
                        else:
                            modalidade_formatada = "Não Informado"

                    # --- 3. N° DO PREGÃO ---
                    num_compra = obter_valor_multiplas_chaves([json_detalhe, licitacao],
                                                              ['numero', 'numeroCompra', 'numero_compra'])
                    ano_compra = obter_valor_multiplas_chaves([json_detalhe, licitacao],
                                                              ['ano', 'anoCompra', 'ano_compra']) or ano

                    if num_compra:
                        numero_pregao = f"{num_compra}/{ano_compra}"
                    else:
                        match = re.search(r'(?:Nº|N°|N\.|N|NÚMERO)\s*(\d+/\d{4})', title, re.IGNORECASE)
                        if match:
                            numero_pregao = match.group(1)
                        else:
                            numero_pregao = str(obter_valor_multiplas_chaves([json_detalhe, licitacao],
                                                                             ['numeroProcesso',
                                                                              'numero_processo']) or f"{seq}/{ano}")

                    # --- 4. UASG ---
                    uasg = str(obter_valor_multiplas_chaves([unidade_orgao, licitacao],
                                                            ['codigoUnidade', 'unidade_codigo',
                                                             'codigo_unidade']) or "N/A")

                    # --- 5. DATAS ---
                    data_fim_str = obter_valor_multiplas_chaves(
                        [json_detalhe, licitacao],
                        [
                            'dataHoraFimRecebimentoPropostas',
                            'dataFimRecebimentoPropostas',
                            'data_fim_recebimento_propostas',
                            'data_hora_fim_recebimento_propostas',
                            'data_fim_vigencia',
                            'dataHoraTerminoRecebimentoPropostas'
                        ]
                    )

                    data_abertura_str = obter_valor_multiplas_chaves(
                        [json_detalhe, licitacao],
                        [
                            'dataHoraAberturaProposta',
                            'dataAberturaProposta',
                            'data_hora_abertura_proposta',
                            'data_abertura_proposta',
                            'data_publicacao_pncp',
                            'dataPublicacaoPncp'
                        ]
                    )

                    data_entrega_formatada = formatar_data(data_fim_str)
                    data_abertura_formatada = formatar_data(data_abertura_str)

                    # --- FILTRO NEGATIVO (RUÍDO) ---
                    palavras_proibidas = [
                        'obra civil', 'pavimentação', 'vigilância armada', 'desarmada',
                        'limpeza e conservação', 'terceirização', 'merenda', 'medicamento',
                        'alimentos', 'veículos', 'locação de veículos', 'asfalto'
                    ]
                    if any(ruido in objeto_lower for ruido in palavras_proibidas):
                        continue

                    # --- VALOR ESTIMADO ---
                    valor_total = obter_valor_multiplas_chaves([json_detalhe, licitacao],
                                                               ['valorTotalEstimado', 'valor_total_estimado',
                                                                'valor_global'])
                    if valor_total:
                        try:
                            valor_formatado = f"R$ {float(valor_total):,.2f}".replace(",", "X").replace(".",
                                                                                                        ",").replace(
                                "X", ".")
                        except ValueError:
                            valor_formatado = str(valor_total)
                    else:
                        valor_formatado = "Não Informado"

                    if baixar_arquivos and cnpj and ano and seq:
                        baixar_arquivos_pncp(str(cnpj), str(ano), str(seq), orgao, scraper, pasta_raiz)

                    resultados_totais.append({
                        'N° do pregão': numero_pregao,
                        'Cliente': orgao,
                        'UASG - Portal': uasg,
                        'Modo de disputa': modalidade_formatada,
                        'Entrega da Proposta': data_entrega_formatada,
                        'Abertura da Proposta': data_abertura_formatada,
                        'Objeto': title if not desc else f"{title}\n{desc}",
                        'Valor estimado': valor_formatado,
                        'link': link_acesso
                    })

                pagina_atual += 1

    df_novos = pd.DataFrame(resultados_totais)
    if not df_novos.empty:
        df_novos = df_novos.drop_duplicates(subset=['link'])

    return df_novos


def salvar_em_excel(df_novos: pd.DataFrame, pasta_raiz: str):
    if df_novos.empty:
        return 0

    nome_arquivo = os.path.join(pasta_raiz, "editais_monitorados.xlsx")

    if os.path.exists(nome_arquivo):
        try:
            df_antigo = pd.read_excel(nome_arquivo)
            df_final = pd.concat([df_antigo, df_novos], ignore_index=True)
            df_final = df_final.drop_duplicates(subset=['link'], keep='last')
        except Exception:
            df_final = df_novos
    else:
        df_final = df_novos

    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        df_final.to_excel(writer, index=False, sheet_name='Oportunidades')

        workbook = writer.book
        worksheet = writer.sheets['Oportunidades']

        header_fill = PatternFill(start_color="3751A3", end_color="3751A3", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True, size=11)
        wrap_alignment = Alignment(wrap_text=True, vertical="top")
        center_alignment = Alignment(horizontal="center", vertical="center")

        larguras = {
            'A': 18,  # N° do pregão
            'B': 45,  # Cliente
            'C': 15,  # UASG
            'D': 35,  # Modo de disputa
            'E': 20,  # Entrega
            'F': 20,  # Abertura
            'G': 65,  # Objeto
            'H': 20,  # Valor
            'I': 45  # link
        }

        for col_idx, col in enumerate(df_final.columns, 1):
            col_letter = get_column_letter(col_idx)

            header_cell = worksheet[f"{col_letter}1"]
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = center_alignment

            if col_letter in larguras:
                worksheet.column_dimensions[col_letter].width = larguras[col_letter]
            else:
                worksheet.column_dimensions[col_letter].width = 30

            for row_idx in range(2, len(df_final) + 2):
                worksheet[f"{col_letter}{row_idx}"].alignment = wrap_alignment

    return len(df_novos)