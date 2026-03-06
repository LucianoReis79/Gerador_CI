import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import io
import os

# 1. Configuração da Página
st.set_page_config(layout="wide", page_title="Gerador de Documentos CI", page_icon="📄")

# Definição dos campos obrigatórios (com underscores para compatibilidade Jinja2)
CAMPOS_SAQUE = [
    "Medicamento", "Rp_Ativo", "Curva_ABC", "CMM", "Validade_Ata", "Estoque", 
    "Estoque_Comprometido", "Estoque_Liq", "Cobertura_Estoque_(dias)", 
    "Cobertura_Prevista", "Nº_Meses", "FATOR_EMB", "Pendencias_de_Entrega", 
    "Outras_aquisicoes_em_andamento", "Pedido_mensal", "CI", "Data", 
    "Previsao_de_ativacao_da_ata", "Valor_pedido", "Tipo_de_demanda", "Programa"
]

CAMPOS_DISPENSA = [
    "Medicamento", "Rp_ativo", "Curva_ABC", "CMM", "Validade_Ata", "Estoque", 
    "Estoque_Comprometido", "Estoque_Liq", "Cobertura_Estoque_(dias)", 
    "Cobertura_Prevista", "Nº_Meses", "FATOR_EMB", "Pendencias_de_Entrega", 
    "Outras_aquisicoes_em_andamento", "Pedido_mensal", "CI", "Data", 
    "Previsao_de_ativacao_da_ata", "Valor_do_pedido", "Tipo_de_demanda", "Programa", 
    "Texto_ajustado", "Texto_ajustado_2"
]

def validar_dados(df, obrigatorios):
    colunas_presentes = set(df.columns)
    return [col for col in obrigatorios if col not in colunas_presentes]

def main():
    st.title("📄 Gerador de Documentos da Programação")
    
    # 2. Seleção de Tipo (Radial) [1]
    tipo_doc = st.radio(
        "Escolha o tipo de documento que deseja gerar:",
        ["Saque RP", "Dispensa"],
        horizontal=True
    )
    
    modelo_path = "saque_rp.docx" if tipo_doc == "Saque RP" else "dispensa.docx"
    campos_necessarios = CAMPOS_SAQUE if tipo_doc == "Saque RP" else CAMPOS_DISPENSA

    # 3. Entrada de Dados
    raw_data = st.text_area("Cole os dados do Excel aqui (inclua o cabeçalho):", height=200)

    if st.button("🚀 Processar Dados"):
        if raw_data.strip():
            try:
                # Converte TAB em DataFrame
                df = pd.read_csv(io.StringIO(raw_data), sep='\t')
                
                # Saneamento de colunas: remove espaços e caracteres especiais para o Jinja2
                df.columns = [c.strip().replace(' ', '_').replace('$', 'R$') for c in df.columns]
                
                faltantes = validar_dados(df, campos_necessarios)
                if faltantes:
                    st.error(f"As seguintes colunas estão faltando ou com nome errado: {', '.join(faltantes)}")
                else:
                    st.session_state['df_gerador'] = df
                    st.success(f"Tabela pronta! {len(df)} registros identificados.")
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao ler dados: {e}")

    # 4. Geração dos Arquivos
    if 'df_gerador' in st.session_state:
        if st.button("🛠️ Gerar Documentos Word"):
            df = st.session_state['df_gerador']
            zip_buffer = io.BytesIO()
            progresso = st.progress(0)
            
            if not os.path.exists(modelo_path):
                st.error(f"Erro: O arquivo '{modelo_path}' não foi encontrado na pasta do projeto.")
                st.stop()

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for idx, row in df.iterrows():
                    doc = DocxTemplate(modelo_path)
                    # Converte linha para dicionário (chaves serão os nomes das colunas)
                    contexto = row.to_dict()
                    doc.render(contexto)
                    
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    
                    # Nome do arquivo baseado no campo CI (limpando caracteres proibidos)
                    nome_ci = str(row['CI']).replace("/", "-").replace("\\", "-")
                    nome_arq = f"{nome_ci}.docx"
                    
                    zip_file.writestr(nome_arq, doc_io.getvalue())
                    progresso.progress((idx + 1) / len(df))
            
            st.session_state['zip_pronto'] = zip_buffer.getvalue()
            st.success(f"Sucesso! {len(df)} documentos gerados no pacote.")

    # 5. Download do ZIP
    if 'zip_pronto' in st.session_state:
        st.download_button(
            label="📥 Baixar Documentos (.ZIP)",
            data=st.session_state['zip_pronto'],
            file_name=f"documentos_{tipo_doc.lower().replace(' ', '_')}.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()