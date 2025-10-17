import streamlit as st
import os
import win32com.client
import time

# --- Configuração da Página e Estilos ---
st.set_page_config(layout="wide", page_title="DocFlow Pro")

# CSS para customizar a aparência e deixar mais parecido com o design
st.markdown("""
<style>
    /* Estilo para os contêineres de categoria */
    .category-box {
        border: 1px solid #e6e6e6;
        border-radius: 0.5rem;
        padding: 1rem;
        margin-bottom: 1rem;
        min-height: 150px;
    }
    .category-box h4 {
        margin-top: 0;
        margin-bottom: 1rem;
    }
    /* Estilo para os itens de arquivo */
    .file-item {
        background-color: #e8f1ff;
        border-radius: 0.25rem;
        padding: 0.5rem;
        margin-bottom: 0.5rem;
        font-family: monospace;
        display: flex;
        align-items: center;
    }
    .file-item::before {
        content: '📄';
        margin-right: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# --- Lógica do App ---

# Inicializa o session_state para guardar as categorias dos arquivos
if 'file_assignments' not in st.session_state:
    st.session_state.file_assignments = {}

# Diretório para arquivos temporários
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'temp_uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- Interface do Usuário (UI) ---

# SIDEBAR (Painel da Direita)
with st.sidebar:
    st.header("Informações do Documento")
    st.caption("Preencha os dados para processar os arquivos")
    
    referencia = st.text_input("Referência", placeholder="REP0395-24")
    numero_po = st.text_input("Número do PO", placeholder="156561")
    referencia_cliente = st.text_input("Referência do Cliente", placeholder="MINE0325-25")
    
    process_button = st.button("Processar e Baixar", type="primary", use_container_width=True)
    # Placeholder para o botão de download
    download_placeholder = st.empty()

# ÁREA PRINCIPAL (Conteúdo da Esquerda)
st.title("DocFlow Pro")
st.markdown("Organize e processe seus documentos")

# Uploader de arquivos
uploaded_files = st.file_uploader(
    "Adicione arquivos para começar",
    type="pdf",
    accept_multiple_files=True,
    label_visibility="collapsed"
)

# Se arquivos foram upados, mostra a interface de categorização
if uploaded_files:
    st.subheader("Categorize seus arquivos")
    
    # Cria colunas para o usuário associar cada arquivo a uma categoria
    for file in uploaded_files:
        # Se um arquivo novo for adicionado, inicializa sua categoria
        if file.file_id not in st.session_state.file_assignments:
            st.session_state.file_assignments[file.file_id] = {"name": file.name, "category": "Outros"}

        category = st.selectbox(
            f"Categoria para **{file.name}**",
            ["Fatura", "Capa de Faturamento", "DI", "Outros"],
            index=["Fatura", "Capa de Faturamento", "DI", "Outros"].index(st.session_state.file_assignments[file.file_id]["category"]),
            key=f"cat_{file.file_id}"
        )
        # Atualiza a categoria do arquivo
        st.session_state.file_assignments[file.file_id]["category"] = category

    st.divider()

    # --- Exibição das Caixas de Categoria ---
    
    # Conta quantos arquivos estão categorizados
    total_files = len(uploaded_files)
    categorized_files = total_files 
    
    st.header(f"Arquivos ({total_files})")
    st.caption(f"{categorized_files} categorizado(s)")

    # Define as categorias para exibição
    categories_to_display = {
        "Fatura": [],
        "Capa de Faturamento": [],
        "DI": [],
        "Outros": []
    }

    # Agrupa os arquivos por categoria
    for file in uploaded_files:
        cat = st.session_state.file_assignments[file.file_id]["category"]
        if cat in categories_to_display:
            categories_to_display[cat].append(file.name)

    # Cria o layout em grade 2x2
    col1, col2 = st.columns(2)

    with col1:
        with st.container():
            st.markdown('<div class="category-box"><h4>Fatura</h4>', unsafe_allow_html=True)
            for file_name in categories_to_display["Fatura"]:
                st.markdown(f'<div class="file-item">{file_name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with st.container():
            st.markdown('<div class="category-box"><h4>DI</h4>', unsafe_allow_html=True)
            for file_name in categories_to_display["DI"]:
                st.markdown(f'<div class="file-item">{file_name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        with st.container():
            st.markdown('<div class="category-box"><h4>Capa de Faturamento</h4>', unsafe_allow_html=True)
            for file_name in categories_to_display["Capa de Faturamento"]:
                st.markdown(f'<div class="file-item">{file_name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with st.container():
            st.markdown('<div class="category-box"><h4>Outros</h4>', unsafe_allow_html=True)
            if not categories_to_display["Outros"] and not categories_to_display["Fatura"] and not categories_to_display["Capa de Faturamento"] and not categories_to_display["DI"]:
                 st.write("Arraste arquivos aqui") # Simula o placeholder
            for file_name in categories_to_display["Outros"]:
                st.markdown(f'<div class="file-item">{file_name}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)


# --- Lógica de Processamento (Backend) ---
if process_button:
    # Validações
    if not all([referencia, numero_po, referencia_cliente]):
        st.sidebar.error("Preencha todos os campos de informação.")
    elif not uploaded_files:
        st.sidebar.error("Nenhum arquivo PDF foi enviado.")
    else:
        with st.sidebar:
            with st.spinner('Processando... Automatizando o Word, por favor aguarde.'):
                word_path = None
                saved_pdf_paths = []
                try:
                    word_app = win32com.client.Dispatch("Word.Application")
                    word_app.Visible = False
                    doc = word_app.Documents.Add()
                    
                    for uploaded_file in uploaded_files:
                        pdf_path = os.path.join(UPLOAD_FOLDER, uploaded_file.name)
                        with open(pdf_path, "wb") as f:
                            f.write(uploaded_file.getbuffer())
                        saved_pdf_paths.append(pdf_path)

                        category = st.session_state.file_assignments[uploaded_file.file_id]["category"]
                        icon_label = f"{referencia}_{numero_po}_{category}.pdf"

                        para = doc.Content.Paragraphs.Add()
                        para.Range.Text = f"Anexo: {icon_label}"
                        para.Range.InsertParagraphAfter()
                        
                        doc.InlineShapes.AddOLEObject(
                            ClassName="AcroExch.Document.DC",
                            FileName=os.path.abspath(pdf_path),
                            LinkToFile=False,
                            DisplayAsIcon=True,
                            IconLabel=icon_label
                        )
                        doc.Content.InsertParagraphAfter()
                    
                    word_filename = f"Processo_{referencia}_{numero_po}.docx"
                    word_path = os.path.join(UPLOAD_FOLDER, word_filename)
                    absolute_word_path = os.path.abspath(word_path)
                    doc.SaveAs(absolute_word_path)
                    st.success(f"Sucesso! Documento gerado.")

                except Exception as e:
                    st.error(f"Erro na automação: {e}")
                
                finally:
                    if 'doc' in locals(): doc.Close(False)
                    if 'word_app' in locals(): word_app.Quit()
                    
                    # Disponibiliza o arquivo para download se foi criado
                    if word_path and os.path.exists(word_path):
                        with open(word_path, "rb") as file_data:
                            download_placeholder.download_button(
                                label="Clique para Baixar o Word",
                                data=file_data,
                                file_name=word_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )
                    
                    # Limpa arquivos temporários
                    for path in saved_pdf_paths:
                        if os.path.exists(path): os.remove(path)
