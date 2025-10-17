import streamlit as st
import os
import win32com.client
import time

# --- Configuração da Página e Diretórios ---
st.set_page_config(layout="wide", page_title="Organizador de Documentos")

# Cria um diretório temporário para uploads na mesma pasta do script
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'temp_uploads')
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# --- Interface do Usuário (UI) com Streamlit ---

st.title("Automatizador de Processos de Documentos")
st.markdown("Faça o upload dos seus PDFs, preencha as informações e gere o documento Word final.")

# 1. Widgets para coletar as informações de referência
st.header("1. Informações do Processo")
referencia = st.text_input("Referência do Processo", placeholder="Ex: REF12345")
numero_po = st.text_input("Número do PO (Purchase Order)", placeholder="Ex: PO67890")
referencia_cliente = st.text_input("Referência do Cliente", placeholder="Ex: CLIENTE-XYZ")

# 2. Widget para upload de múltiplos arquivos
st.header("2. Upload dos Arquivos PDF")
uploaded_files = st.file_uploader(
    "Arraste e solte os arquivos aqui",
    type="pdf",
    accept_multiple_files=True
)

# 3. Widget para categorizar cada arquivo
st.header("3. Categorize os Documentos")
if uploaded_files:
    file_categories = {}
    for uploaded_file in uploaded_files:
        # Cria uma linha com o nome do arquivo e uma caixa de seleção ao lado
        col1, col2 = st.columns([3, 2])
        with col1:
            st.write(f"📄 {uploaded_file.name}")
        with col2:
            # A caixa de seleção é a "caixinha" da sua ideia original
            category = st.selectbox(
                f"Tipo para {uploaded_file.name}",
                ["", "Fatura", "Capa de Faturamento", "DI", "BL", "AWB", "Outro"],
                key=uploaded_file.name  # A chave única é essencial aqui
            )
            file_categories[uploaded_file.name] = category
else:
    st.info("Aguardando o upload dos arquivos para categorização.")


# 4. Botão para iniciar o processamento
st.header("4. Gerar Documento Final")
if st.button("Processar e Gerar Word"):
    # Validações básicas
    if not all([referencia, numero_po]):
        st.error("Por favor, preencha a Referência do Processo e o Número do PO.")
    elif not uploaded_files:
        st.error("Nenhum arquivo PDF foi enviado.")
    elif any(cat == "" for cat in file_categories.values()):
        st.warning("Atenção: Um ou mais arquivos não foram categorizados. Eles serão incluídos com a categoria 'SemCategoria'.")
    else:
        # Se tudo estiver OK, começa o processamento
        with st.spinner('Automatizando o Word... Este processo pode levar um momento.'):
            try:
                word_app = win32com.client.Dispatch("Word.Application")
                word_app.Visible = False
                doc = word_app.Documents.Add()
                
                saved_pdf_paths = []
                
                for uploaded_file in uploaded_files:
                    # Salva o arquivo enviado em um local temporário
                    pdf_path = os.path.join(UPLOAD_FOLDER, uploaded_file.name)
                    with open(pdf_path, "wb") as f:
                        f.write(uploaded_file.getbuffer())
                    saved_pdf_paths.append(pdf_path)

                    category = file_categories.get(uploaded_file.name, "SemCategoria")
                    icon_label = f"{referencia}_{numero_po}_{category}.pdf"

                    # Adiciona texto e o objeto PDF incorporado
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
                
                # Salva o documento Word
                word_filename = f"Processo_{referencia}_{numero_po}.docx"
                word_path = os.path.join(UPLOAD_FOLDER, word_filename)
                absolute_word_path = os.path.abspath(word_path)
                doc.SaveAs(absolute_word_path)
                
                st.success(f"Documento '{word_filename}' gerado com sucesso!")

                # Disponibiliza o arquivo para download
                with open(absolute_word_path, "rb") as file_data:
                    st.download_button(
                        label="Clique aqui para baixar o Word",
                        data=file_data,
                        file_name=word_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"Ocorreu um erro durante a automação do Word: {e}")
            
            finally:
                # Garante que o Word seja fechado e os arquivos limpos
                if 'doc' in locals():
                    doc.Close(False)
                if 'word_app' in locals():
                    word_app.Quit()
                
                for path in saved_pdf_paths:
                    if os.path.exists(path):
                        os.remove(path)
