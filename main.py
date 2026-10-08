import cloudscraper
import pandas as pd
import os
import time
import requests
import re
from datetime import datetime


def baixar_arquivos_pncp(cnpj: str, ano: str, seq: str, orgao: str, scraper):
    """Acessa a API de anexos do PNCP, rastreia ou constrói o link correto e baixa os arquivos."""
    pasta_raiz = "Editais_Baixados"

    # Limpeza do nome do órgão para criar a pasta
    orgao_limpo = re.sub(r'[\\/*?:"<>|]', "", str(orgao)).strip()[:60]
    if not orgao_limpo:
        orgao_limpo = f"Orgao_{cnpj}"

    nome_pasta = f"{orgao_limpo}_Seq_{seq}"
    pasta_destino = os.path.join(pasta_raiz, nome_pasta)

    url_lista_arquivos = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}/arquivos"

    try:
        res = scraper.get(url_lista_arquivos, headers={'Accept': 'application/json'}, timeout=15)
        if res.status_code != 200:
            return

        arquivos = res.json()
        if not arquivos:
            return

        os.makedirs(pasta_destino, exist_ok=True)
        print(f"       Detectados {len(arquivos)} arquivo(s). Baixando na pasta '{nome_pasta}'...")

        for arq in arquivos:
            titulo = arq.get('tituloDocumento') or arq.get('titulo') or "Documento"

            # 1. TENTA CAPTURAR A URL PELAS CHAVES POSSÍVEIS
            url_down = arq.get('url') or arq.get('linkArquivo') or arq.get('urlDocumento')

            # 2. SE NÃO VEIO LINK, CONSTRÓI A URL MATEMATICAMENTE
            if not url_down:
                seq_arq = arq.get('sequencialDocumento') or arq.get('sequencialArquivo') or arq.get('sequencial')
                if seq_arq:
                    url_down = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}/arquivos/{seq_arq}"

            # Se mesmo assim não houver como baixar, avisa
            if not url_down:
                print(f"      [!] O portal ocultou completamente o acesso para '{titulo}'.")
                continue

            # Limpa o título do arquivo
            titulo_limpo = re.sub(r'[\\/*?:"<>|]', "", str(titulo)).strip()[:80]
            if not titulo_limpo:
                titulo_limpo = f"Documento_{seq_arq}"

            if not '.' in titulo_limpo[-5:]:
                titulo_limpo += ".pdf"

            caminho_arquivo = os.path.join(pasta_destino, titulo_limpo)

            if os.path.exists(caminho_arquivo):
                continue

            try:
                # 3. FAZ O DOWNLOAD DO ARQUIVO BINÁRIO
                doc_res = scraper.get(url_down, stream=True, allow_redirects=True, timeout=60)

                if doc_res.status_code == 200:
                    content_type = doc_res.headers.get('Content-Type', '')

                    # Se o governo retornou um JSON em vez de um arquivo, o documento está corrompido ou bloqueado lá
                    if 'application/json' in content_type:
                        msg = doc_res.json()
                        print(f"      [!] O servidor recusou o envio de '{titulo_limpo}'. Retorno: {msg}")
                        continue

                    # Escreve o arquivo na pasta
                    with open(caminho_arquivo, 'wb') as f:
                        for chunk in doc_res.iter_content(chunk_size=8192):
                            if chunk:
                                f.write(chunk)
                    print(f"        -> '{titulo_limpo}' salvo com sucesso!")
                else:
                    print(
                        f"      [!] Falha no servidor do governo ao acessar '{titulo_limpo}' (Erro HTTP: {doc_res.status_code})")
            except Exception as e:
                print(f"      [!] Erro de conexão ao baixar '{titulo_limpo}': {e}")

    except Exception as e:
        print(f"      [!] Erro ao tentar acessar a lista de anexos: {e}")


def buscar_editais_pncp(palavras_chave: list) -> pd.DataFrame:
    base_url = "https://pncp.gov.br/api/search"
    resultados_totais = []
    hoje = datetime.now()

    scraper = cloudscraper.create_scraper(browser={'browser': 'chrome', 'platform': 'windows', 'desktop': True})
    headers_detalhe = {'Accept': 'application/json'}

    print(f"Iniciando varredura no PNCP... (Data base do sistema: {hoje.strftime('%d/%m/%Y %H:%M')})")

    status_abertos = ['recebendo_proposta']

    for keyword in palavras_chave:
        for status_slug in status_abertos:
            params = {
                'q': keyword,
                'tipos_documento': 'edital',
                'status': status_slug,
                'ordenacao': '-data',
                'pagina': 1,
                'tam_pagina': 50
            }

            max_tentativas = 3
            for tentativa in range(1, max_tentativas + 1):
                try:
                    print(f"\n-> Pesquisando por: '{keyword}' | Filtro: Recebendo Propostas")

                    response = scraper.get(base_url, params=params, timeout=60)
                    response.raise_for_status()

                    dados = response.json().get('items', [])

                    for licitacao in dados:
                        id_licitacao = licitacao.get('numero_controle_pncp', 'N/A')
                        cnpj = licitacao.get('orgao_cnpj')
                        ano = licitacao.get('ano')
                        seq = licitacao.get('numero_sequencial')

                        data_fim_str = licitacao.get('data_fim_recebimento_propostas')

                        if (not data_fim_str or str(data_fim_str).strip() == "") and cnpj and ano and seq:
                            detalhe_url = f"https://pncp.gov.br/api/pncp/v1/orgaos/{cnpj}/compras/{ano}/{seq}"
                            try:
                                res_detalhe = scraper.get(detalhe_url, headers=headers_detalhe, timeout=15)
                                if res_detalhe.status_code == 200:
                                    json_detalhe = res_detalhe.json()
                                    data_fim_str = (
                                            json_detalhe.get('dataHoraFimRecebimentoPropostas') or
                                            json_detalhe.get('dataFimRecebimentoPropostas') or
                                            json_detalhe.get('dataHoraAberturaProposta') or
                                            json_detalhe.get('dataAberturaProposta')
                                    )
                            except Exception:
                                pass

                        is_valid = True
                        data_fim_formatada = "Verificar no Edital"

                        if data_fim_str and str(data_fim_str).strip() != "":
                            try:
                                data_fim_dt = pd.to_datetime(data_fim_str).tz_localize(None)

                                if data_fim_dt < hoje:
                                    print(
                                        f"   [DESCARTADO] {id_licitacao} - Prazo encerrou em: {data_fim_dt.strftime('%d/%m/%Y %H:%M')}")
                                    is_valid = False
                                else:
                                    data_fim_formatada = data_fim_dt.strftime('%d/%m/%Y %H:%M')
                            except:
                                pass

                        if not is_valid:
                            continue

                        orgao = licitacao.get('orgao_nome', 'Não Informado')

                        print(f"   [ADICIONADO] {id_licitacao} - Válido até: {data_fim_formatada}")

                        # ---> MÓDULO DE DOWNLOAD ENTRA AQUI <---
                        if cnpj and ano and seq:
                            baixar_arquivos_pncp(str(cnpj), str(ano), str(seq), orgao, scraper)

                        objeto_texto = str(
                            licitacao.get('description') or
                            licitacao.get('title') or
                            licitacao.get('informacao_complementar') or
                            ""
                        ).strip()

                        data_pub_str = licitacao.get('data_publicacao_pncp')
                        data_pub_formatada = pd.to_datetime(data_pub_str.split('T')[0]).strftime(
                            '%d/%m/%Y') if data_pub_str else 'N/A'

                        if cnpj and ano and seq:
                            link_acesso = f"https://pncp.gov.br/app/editais/{cnpj}/{ano}/{seq}"
                        else:
                            link_acesso = licitacao.get('itemUrl', 'Link indisponível')

                        resultados_totais.append({
                            'id_licitacao': id_licitacao,
                            'orgao_nome': orgao,
                            'objeto': objeto_texto,
                            'data_publicacao': data_pub_formatada,
                            'data_limite_proposta': data_fim_formatada,
                            'status_pncp': "Recebendo Propostas",
                            'link_acesso': link_acesso
                        })

                    break

                except requests.exceptions.Timeout:
                    print("   [!] O servidor demorou muito para responder.")
                    if tentativa < max_tentativas:
                        time.sleep(5)

                except Exception as e:
                    print(f"   [!] Erro de conexão: {e}")
                    if tentativa < max_tentativas:
                        time.sleep(5)

    df = pd.DataFrame(resultados_totais)

    if not df.empty:
        df = df.drop_duplicates(subset=['id_licitacao'])

    return df


def salvar_em_csv(df_editais: pd.DataFrame, nome_arquivo="editais_monitorados.csv"):
    if df_editais.empty:
        print("\nNenhum edital ativo encontrado nesta varredura.")
        return

    df_editais.to_csv(nome_arquivo, sep=';', index=False, encoding='utf-8-sig')
    print(f"\n[SUCESSO] Arquivo '{nome_arquivo}' salvo com base atualizada! Total: {len(df_editais)}")


if __name__ == "__main__":
    print("--- SISTEMA DE INTELIGÊNCIA E MONITORAMENTO DE EDITAIS ---")

    termos_busca = ['videowall']

    df_novos = buscar_editais_pncp(termos_busca)
    salvar_em_csv(df_novos)