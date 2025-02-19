import pandas as pd
import requests
import time
import random
import streamlit as st
import io

# Função para limpar o CNPJ (remover caracteres indesejados)
def limpar_cnpj(cnpj):
    cnpj = str(cnpj).strip().replace(".", "").replace("-", "").replace("/", "")
    return cnpj.zfill(14)  # Garante que tenha 14 dígitos

# Função para consultar o CNPJ
def consultar_cnpj(cnpj):
    cnpj = limpar_cnpj(cnpj)
    url = f"https://www.receitaws.com.br/v1/cnpj/{cnpj}"

    while True:
        try:
            response = requests.get(url)
            if response.status_code == 200:
                dados_cnpj = response.json()
                if 'erro' in dados_cnpj:
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
                st.warning(f"🚨 API bloqueada! Aguardando antes de tentar novamente...")
                time.sleep(random.uniform(15, 30))

            else:
                return None

        except requests.exceptions.RequestException:
            return None

# Função para processar os CNPJs da planilha e gerar o DataFrame
def processar_cnpjs(arquivo_excel):
    df = pd.read_excel(arquivo_excel, dtype=str)
    
    if "CNPJ" not in df.columns:
        st.error("Erro: A planilha deve conter uma coluna chamada 'CNPJ'.")
        return None

    dados = []
    cnpjs_pendentes = df["CNPJ"].dropna().apply(limpar_cnpj).tolist()

    while cnpjs_pendentes:
        st.write(f"🔄 Consultando {len(cnpjs_pendentes)} CNPJs...")
        novos_dados = []
        cnpjs_falha = []

        for cnpj in cnpjs_pendentes:
            resultado = consultar_cnpj(cnpj)
            if resultado:
                novos_dados.append(resultado)
            else:
                cnpjs_falha.append(cnpj)

            time.sleep(random.uniform(3, 6))  # Evita bloqueio da API

        dados.extend(novos_dados)
        cnpjs_pendentes = cnpjs_falha

        if cnpjs_pendentes:
            st.warning(f"⏳ {len(cnpjs_pendentes)} CNPJs ainda falharam. Tentando novamente em 60 segundos...")
            time.sleep(60)

    df_resultado = pd.DataFrame(dados)
    return df_resultado

# Função para permitir download do Excel no Streamlit
def download_planilha(df_resultado):
    if df_resultado is None or df_resultado.empty:
        st.error("Nenhum dado para salvar.")
        return

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_resultado.to_excel(writer, sheet_name="CNPJs", index=False)
    
    st.download_button(
        label="📥 Baixar Planilha",
        data=output.getvalue(),
        file_name="dados_cnpjs.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# Interface Streamlit
st.title("🔍 Consulta de CNPJs")
st.write("Faça o upload de uma planilha com a coluna 'CNPJ' para realizar a consulta.")

# Widget para upload do arquivo Excel
arquivo_upload = st.file_uploader("📂 Escolha um arquivo Excel", type=['xlsx'])

if arquivo_upload is not None:
    df_resultado = processar_cnpjs(arquivo_upload)

    if df_resultado is not None:
        st.write(df_resultado)
        download_planilha(df_resultado)
