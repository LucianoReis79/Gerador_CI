import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import zipfile
import io
import os

# 1. Configuração da Página [1]
st.set_page_config(layout="wide", page_title="Gerador de Documentos CI", page_icon="📄")

# Definição dos campos obrigatórios (extraídos de app_2026.txt) [1, 2]
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
    # 2. Instruções na Barra Lateral (Bloco Azul Estilizado) [3]
    instrucoes_html = """
    <div style="background-color: #D1E8FF; padding: 15px; border-radius: 5px; border-left: 8px solid #003366; color: #000; margin-bottom: 20px;">
        <h3 style="margin-top: 0; font-size: 18px;">📋 Instruções de Uso</h3>
        <ul style="font-size: 13px; line-height: 1.5; padding-left: 20px;">
            <li><b>Navegação:</b> Escolha o tipo de documento no seletor.</li>
            <li><b>Copiar Dados:</b> No Excel, selecione as linhas com o <b>cabeçalho</b> e use Ctrl+C.</li>
            <li><b>Processamento:</b> Cole os dados e clique em <b>🚀 Processar Dados</b>.</li>
            <li><b>Geração:</b> Clique em <b>🛠️ Gerar Documentos</b>.</li>
            <li><b>Download:</b> Clique em <b>📥 Baixar Documentos (.ZIP)</b>.</li>
        </ul>
        
    </div>
    """
    st.sidebar.markdown(instrucoes_html, unsafe_allow_html=True)
    st.sidebar.markdown("---")
    st.sidebar.caption("⭐ Desenvolvido por Luciano Reis")

    st.title("📄 Gerador de Documentos da Programação")

    # 3. CSS para fundo azul nos botões de escolha (radio) [4]
    st.markdown("""
        <style>
        div[data-testid="stRadio"] > div {
            background-color: #D1E8FF;
            padding: 15px;
            border-radius: 10px;
        }
        </style>
    """, unsafe_allow_html=True)

    tipo_doc = st.radio(
        "Escolha o tipo de documento que deseja gerar:",
        ["Saque RP", "Dispensa"],
        horizontal=True
    )
    
    modelo_path = "saque_rp.docx" if tipo_doc == "Saque RP" else "dispensa.docx"
    campos_necessarios = CAMPOS_SAQUE if tipo_doc == "Saque RP" else CAMPOS_DISPENSA

    # 4. Entrada de Dados com rótulo em negrito
    raw_data = st.text_area("**Cole os dados do Excel aqui (inclua o cabeçalho):**", height=200)

    if st.button("🚀 Processar Dados"):
        if raw_data.strip():
            try:
                # Converte TAB em DataFrame e saneia colunas para Jinja2
                df = pd.read_csv(io.StringIO(raw_data), sep='\t')
                df.columns = [c.strip().replace(' ', '_').replace('$', 'R$') for c in df.columns]
                
                faltantes = validar_dados(df, campos_necessarios)
                if faltantes:
                    st.error(f"As seguintes colunas estão faltando ou com nome errado: {', '.join(faltantes)}")
                else:
                    st.session_state['df_gerador'] = df [5]
                    st.success(f"Tabela pronta! {len(df)} registros identificados.")
                    st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao ler dados: {e}")

    # 5. Geração dos Arquivos Word baseados no campo CI
    if 'df_gerador' in st.session_state:
        if st.button("🛠️ Gerar Documentos Word"):
            df = st.session_state['df_gerador']
            zip_buffer = io.BytesIO()
            progresso = st.progress(0)
            
            if not os.path.exists(modelo_path):
                st.error(f"Erro: O arquivo '{modelo_path}' não foi encontrado.")
                st.stop()

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                for idx, row in df.iterrows():
                    doc = DocxTemplate(modelo_path)
                    doc.render(row.to_dict())
                    
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    
                    # Nome do arquivo baseado na coluna CI
                    nome_ci = str(row['CI']).replace("/", "-").replace("\\", "-")
                    zip_file.writestr(f"{nome_ci}.docx", doc_io.getvalue())
                    progresso.progress((idx + 1) / len(df))
            
            st.session_state['zip_pronto'] = zip_buffer.getvalue() [6]
            st.success("Sucesso! Documentos gerados no pacote ZIP.")

    # 6. Botão de Download [7]
    if 'zip_pronto' in st.session_state:
        st.download_button(
            label="📥 Baixar Documentos (.ZIP)",
            data=st.session_state['zip_pronto'],
            file_name=f"documentos_{tipo_doc.lower().replace(' ', '_')}.zip",
            mime="application/zip"
        )

if __name__ == "__main__":
    main() # [3]