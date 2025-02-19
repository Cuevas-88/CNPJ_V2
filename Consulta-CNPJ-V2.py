import pandas as pd
import requests
import time
import random
import os

# Função para limpar o CNPJ (remover caracteres indesejados)
def limpar_cnpj(cnpj):
    cnpj = str(cnpj).strip().replace(".", "").replace("-", "").replace("/", "")
    return cnpj.zfill(14)  # Garante que tenha 14 dígitos

# Função para consultar o CNPJ
def consultar_cnpj(cnpj):
    cnpj = limpar_cnpj(cnpj)

    if len(cnpj) != 14:
        print(f"CNPJ inválido: {cnpj}")
        return None

    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"

    while True:  # Loop até conseguir resposta válida
        try:
            response = requests.get(url)

            if response.status_code == 200:
                dados_cnpj = response.json()

                if 'erro' in dados_cnpj:
                    print(f"Erro na API para o CNPJ {cnpj}: {dados_cnpj['erro']}")
                    return None

                # Capturando Atividade Principal
                atividade_principal_codigo = dados_cnpj['atividade_principal'][0].get('code', '') if 'atividade_principal' in dados_cnpj and dados_cnpj['atividade_principal'] else ''
                atividade_principal_descricao = dados_cnpj['atividade_principal'][0].get('text', '') if 'atividade_principal' in dados_cnpj and dados_cnpj['atividade_principal'] else ''

                # Capturando todas as Atividades Secundárias
                cnaes_secundarios_codigos = [atividade.get('code', '') for atividade in dados_cnpj.get('atividades_secundarias', [])]
                cnaes_secundarios_descricoes = [atividade.get('text', '') for atividade in dados_cnpj.get('atividades_secundarias', [])]

                # Capturando todo o Quadro Societário (QSA)
                quadro_societario = [
                    f"{socio.get('nome', 'Não disponível')} ({socio.get('qual', 'Não informado')})"
                    for socio in dados_cnpj.get('qsa', [])
                ]
                qsa_formatado = "; ".join(quadro_societario) if quadro_societario else "Não disponível"

                return {
                    'CNPJ': dados_cnpj.get('cnpj', ''),
                    'Nome': dados_cnpj.get('nome', ''),
                    'Nome Fantasia': dados_cnpj.get('fantasia', ''),
                    'Natureza Jurídica': dados_cnpj.get('natureza_juridica', ''),
                    'Endereço': f"{dados_cnpj.get('logradouro', '')}, {dados_cnpj.get('numero', '')} - {dados_cnpj.get('bairro', '')}, {dados_cnpj.get('municipio', '')} - {dados_cnpj.get('uf', '')}",
                    'Telefone': dados_cnpj.get('telefone', ''),
                    'Email': dados_cnpj.get('email', ''),
                    'Atividade Principal (CNAE)': atividade_principal_codigo,
                    'Descrição Atividade Principal': atividade_principal_descricao,
                    'CNAEs Secundários (Códigos)': ", ".join(cnaes_secundarios_codigos),
                    'CNAEs Secundários (Descrições)': ", ".join(cnaes_secundarios_descricoes),
                    'Situação Cadastral': dados_cnpj.get('situacao', ''),
                    'Data de Abertura': dados_cnpj.get('abertura', ''),
                    'QSA (Sócios e Administradores)': qsa_formatado
                }

            elif response.status_code == 429:
                print(f"🚨 API bloqueada (Erro 429)! Aguardando antes de tentar novamente...")
                time.sleep(random.uniform(15, 30))  # Espera entre 15 e 30 segundos

            else:
                print(f"Erro HTTP {response.status_code} para o CNPJ {cnpj}")
                return None

        except requests.exceptions.RequestException as e:
            print(f"Erro de conexão para o CNPJ {cnpj}: {e}")
            return None

# Função para processar os CNPJs da planilha e gerar o DataFrame
def processar_cnpjs(nome_arquivo):
    df = pd.read_excel(nome_arquivo, dtype=str)  # Carregar como string para evitar problemas de formatação
    
    if "CNPJ" not in df.columns:
        print("Erro: A planilha deve conter uma coluna chamada 'CNPJ'.")
        return None

    dados = []
    cnpjs_pendentes = df["CNPJ"].dropna().apply(limpar_cnpj).tolist()  # Lista inicial de CNPJs a consultar

    while cnpjs_pendentes:  # Continua até não haver mais CNPJs pendentes
        print(f"🔄 Consultando {len(cnpjs_pendentes)} CNPJs...")
        novos_dados = []
        cnpjs_falha = []

        for cnpj in cnpjs_pendentes:
            resultado = consultar_cnpj(cnpj)
            if resultado:
                novos_dados.append(resultado)
            else:
                cnpjs_falha.append(cnpj)  # Guarda apenas os que falharam

            time.sleep(random.uniform(3, 6))  # Evita bloqueio da API

        dados.extend(novos_dados)  # Adiciona os novos resultados à lista principal
        cnpjs_pendentes = cnpjs_falha  # Atualiza a lista com os CNPJs que ainda precisam ser consultados

        if cnpjs_pendentes:
            print(f"⏳ {len(cnpjs_pendentes)} CNPJs ainda falharam. Tentando novamente em 60 segundos...")
            time.sleep(60)  # Aguarda antes de tentar novamente

    df_resultado = pd.DataFrame(dados)
    return df_resultado

# Função para salvar e baixar o arquivo final
def salvar_planilha(df_resultado, nome_arquivo_saida="dados_cnpjs.xlsx"):
    if df_resultado is None or df_resultado.empty:
        print("Nenhum dado para salvar.")
        return
    
    df_resultado.to_excel(nome_arquivo_saida, index=False)
    print(f"✅ Planilha salva com sucesso: {os.path.abspath(nome_arquivo_saida)}")

# 🚀 Passo 1: Solicitar o nome do arquivo de entrada
nome_arquivo_entrada = input("Digite o nome do arquivo Excel com os CNPJs (ex: cnpjs.xlsx): ")

if not os.path.exists(nome_arquivo_entrada):
    print("❌ Arquivo não encontrado! Verifique o nome e tente novamente.")
else:
    # 🚀 Passo 2: Processar o arquivo
    df_resultado = processar_cnpjs(nome_arquivo_entrada)

    if df_resultado is not None:
        # 🚀 Passo 3: Salvar os resultados
        salvar_planilha(df_resultado)
        print("🎯 Consulta finalizada com sucesso!")
