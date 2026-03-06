Gerador de Documentos de Programação de Medicamentos
Este aplicativo web automatiza a geração de documentos Word (.docx) padronizados a partir de dados administrativos copiados diretamente de planilhas Excel [user prompt]. A ferramenta visa eliminar o preenchimento manual, reduzindo erros e otimizando o tempo em ambientes administrativos [user prompt].
🚀 Tecnologias Utilizadas
Python: Linguagem base para o desenvolvimento [user prompt, 436].
Streamlit: Framework para criação da interface web interativa [user prompt, 439].
Pandas: Biblioteca para manipulação e tratamento de dados em tabelas (DataFrames) [user prompt, 403, 440].
DocxTpl: Utilizado para preencher modelos Word usando tags Jinja2 [user prompt].
OpenPyXL: Garante a compatibilidade na leitura de arquivos modernos do Excel (.xlsx)
.
✨ Funcionalidades
Entrada de Dados via Clipboard: Permite colar dados do Excel (separados por TAB) diretamente na interface [user prompt].
Múltiplos Modelos: Suporte para geração de documentos do tipo Saque RP e Dispensa [user prompt, 460].
Validação Automática: O sistema verifica a presença de cabeçalhos e campos obrigatórios antes do processamento [user prompt].
Geração em Lote: Cria arquivos individuais nomeados automaticamente com base no campo CI de cada registro [user prompt].
Download em ZIP: Agrupa todos os documentos gerados em um único arquivo compactado para facilitar o download [user prompt].
📦 Instalação
Para rodar o projeto localmente, instale as dependências necessárias através do terminal [user prompt, 439, 440]:
pip install streamlit pandas docxtpl python-docx openpyxl
🛠️ Como Executar
Na pasta do projeto, inicie a aplicação com o comando
:
streamlit run app.py
O servidor local será iniciado e o aplicativo abrirá automaticamente em seu navegador padrão
.

--------------------------------------------------------------------------------
Desenvolvido por: Luciano Reis
.
