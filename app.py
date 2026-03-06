import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import io
import os

# Configuração da página profissional [1, 4]
st.set_page_config(layout="wide", page_title="Gerador de Documentos", page_icon="📄")

# Estilos de campos obrigatórios [Turnos anteriores]
CAMPOS_TIPO_1 = [
    "Medicamento", "Rp Ativo", "Curva ABC", "CMM", "Validade Ata", "Estoque", 
    "Estoque Comprometido", "Estoque Liq", "Cobertura Estoque (dias)", 
    "Cobertura Prevista", "Nº Meses", "FATOR_EMB", "Pendencias de Entrega", 
    "Outras aquisições em andamento", "Pedido mensal", "CI", "Data", 
    "Previsão de ativação da ata", "Valor pedido R$", "Tipo de demanda", "Programa"
]

CAMPOS_TIPO_2 = [
    "Medicamento", "Rp ativo", "Curva ABC", "CMM", "Validade Ata", "Estoque", 
    "Estoque Comprometido", "Estoque Liq", "Cobertura Estoque (dias)", 
    "Cobertura Prevista", "Nº Meses", "FATOR_EMB", "Pendencias de Entrega", 
    "Outras aquisições em andamento", "Pedido mensal", "CI", "Data", 
    "Previsão de ativação da ata", "Valor do pedido", "Tipo", "Programa", 
    "Texto ajustado", "Texto ajustado 2"
]

def validar_dados(df, obrigatorios):
    """Verifica se todas as colunas necessárias existem [1, 5]"""
    colunas_presentes = set(df.columns)
    faltantes = [col for col in obrigatorios if col not in colunas_presentes]
    return faltantes

def main():
    st.title("📄 Gerador de Documentos de Programação de Medicamentos")
    
    # 1. Interface de Seleção [6, 7]
    tipo_doc = st.selectbox("Selecione o Tipo de Documento:", ["Documento Tipo 1", "Documento Tipo 2"])
    modelo_path = "modelo1.docx" if tipo_doc == "Documento Tipo 1" else "modelo2.docx"
    campos_necessarios = CAMPOS_TIPO_1 if tipo_doc == "Documento Tipo 1" else CAMPOS_TIPO_2

    # 2. Área de Texto para Dados [8]
    raw_data = st.text_area("Cole os dados copiados do Excel aqui (incluindo cabeçalhos):", height=200)

    if st.button("🚀 Processar dados"):
        if raw_data.strip():
            try:
                # Converte texto (TAB do Excel) em DataFrame [2, 9]
                df = pd.read_csv(io.StringIO(raw_data), sep='\t')
                df.columns = df.columns.str.strip() # Limpeza [1]
                
                # Validação
                faltantes = validar_dados(df, campos_necessarios)
                if faltantes:
                    st.error(f"Erro: As colunas a seguir não foram encontradas: {', '.join(faltantes)}")
                else:
                    st.session_state['df_gerador'] = df
                    st.success(f"Tabela processada com {len(df)} linhas!")
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao interpretar os dados: {e}")
        else:
            st.warning("A área de texto está vazia.")

    # 3. Geração de Documentos [10, 11]
    if 'df_gerador' in st.session_state:
        if st.button("🛠️ Gerar documentos"):
            df = st.session_state['df_gerador']
            zip_buffer = io.BytesIO()
            progresso = st.progress(0)
            contador = 0
            
            # Verifica se modelo existe [12]
            if not os.path.exists(modelo_path):
                st.error(f"Modelo '{modelo_path}' não encontrado na pasta.")
                st.stop()

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for idx, row in df.iterrows():
                    # Preenche o Word via DocxTpl
                    doc = DocxTemplate(modelo_path)
                    contexto = row.to_dict()
                    doc.render(contexto)
                    
                    # Salva no buffer em memória
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    
                    # Nome do arquivo sanitizado
                    nome_arq = f"doc_{idx+1}_{str(row['Medicamento'])[:20]}.docx".replace("/", "-")
                    zip_file.writestr(nome_arq, doc_io.getvalue())
                    
                    # Atualiza progresso [13]
                    contador += 1
                    progresso.progress(contador / len(df))
            
            st.session_state['zip_pronto'] = zip_buffer.getvalue()
            st.success(f"Concluído! {contador} documentos gerados.")

    # 4. Download [14]
    if 'zip_pronto' in st.session_state:
        st.download_button(
            label="📥 Baixar documentos.zip",
            data=st.session_state['zip_pronto'],
            file_name="documentos.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main()
