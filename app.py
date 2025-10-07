"""
Sistema RAG con conexión a SharePoint
Procesa múltiples carpetas y documentos de SharePoint automáticamente
"""

import streamlit as st
import os
import tempfile
from pathlib import Path
import io

# Verificar e importar dependencias
try:
    from langchain.text_splitter import RecursiveCharacterTextSplitter
    from langchain.chains import RetrievalQA
    from langchain.prompts import PromptTemplate
    from langchain.schema import Document
    
    # Importaciones community
    from langchain_community.document_loaders import PyPDFLoader, TextLoader, Docx2txtLoader, UnstructuredPowerPointLoader
    from langchain_community.embeddings import HuggingFaceEmbeddings
    from langchain_community.vectorstores import FAISS
    from langchain_community.llms import Ollama
    
    # SharePoint
    from office365.runtime.auth.authentication_context import AuthenticationContext
    from office365.sharepoint.client_context import ClientContext
    from office365.sharepoint.files.file import File
    
    DEPENDENCIAS_OK = True
except ImportError as e:
    DEPENDENCIAS_OK = False
    ERROR_IMPORT = str(e)

# Configuración
st.set_page_config(
    page_title="RAG SharePoint",
    page_icon="📁",
    layout="wide"
)

# Verificar dependencias primero
if not DEPENDENCIAS_OK:
    st.error(f"❌ Error de dependencias: {ERROR_IMPORT}")
    st.markdown("""
    ### 📦 Instala las dependencias necesarias:
    
    ```bash
    pip install langchain langchain-community faiss-cpu sentence-transformers 
    pip install pypdf python-docx docx2txt unstructured python-pptx
    pip install Office365-REST-Python-Client
    ```
    
    ### 📄 O crea un archivo `requirements.txt`:
    ```
    streamlit
    langchain
    langchain-community
    faiss-cpu
    sentence-transformers
    pypdf
    python-docx
    docx2txt
    unstructured
    python-pptx
    Office365-REST-Python-Client
    ```
    
    Y ejecuta: `pip install -r requirements.txt`
    """)
    st.stop()

# Session state
if 'vectorstore' not in st.session_state:
    st.session_state.vectorstore = None
if 'qa_chain' not in st.session_state:
    st.session_state.qa_chain = None
if 'historial' not in st.session_state:
    st.session_state.historial = []
if 'documentos_cargados' not in st.session_state:
    st.session_state.documentos_cargados = []
if 'sharepoint_conectado' not in st.session_state:
    st.session_state.sharepoint_conectado = False

@st.cache_resource
def inicializar_embeddings():
    """Inicializa embeddings multilingües"""
    return HuggingFaceEmbeddings(
        model_name="sentence-transformers/paraphrase-multilingual-MiniLM-L12-v2",
        model_kwargs={'device': 'cpu'}
    )

@st.cache_resource
def inicializar_llm(modelo="llama2"):
    """Inicializa LLM - cambiar a OpenAI para producción"""
    # Para producción:
    # from langchain_openai import ChatOpenAI
    # return ChatOpenAI(model="gpt-3.5-turbo", api_key=st.secrets["OPENAI_API_KEY"])
    return Ollama(model=modelo, temperature=0.7)

def conectar_sharepoint(site_url, username, password):
    """
    Conecta a SharePoint usando credenciales
    
    Args:
        site_url: URL del sitio SharePoint (ej: https://miempresa.sharepoint.com/sites/misite)
        username: Email corporativo
        password: Contraseña
    """
    try:
        ctx_auth = AuthenticationContext(site_url)
        if ctx_auth.acquire_token_for_user(username, password):
            ctx = ClientContext(site_url, ctx_auth)
            return ctx
        else:
            st.error("❌ Error de autenticación")
            return None
    except Exception as e:
        st.error(f"❌ Error al conectar: {str(e)}")
        return None

def listar_carpetas_recursivo(ctx, folder_url, carpetas_list, nivel=0):
    """Lista todas las carpetas de forma recursiva"""
    try:
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        ctx.load(folder)
        ctx.execute_query()
        
        # Obtener subcarpetas
        folders = folder.folders
        ctx.load(folders)
        ctx.execute_query()
        
        for subfolder in folders:
            # Ignorar carpetas del sistema
            if subfolder.properties['Name'] not in ['Forms', 'Item']:
                carpeta_info = {
                    'nombre': subfolder.properties['Name'],
                    'url': subfolder.properties['ServerRelativeUrl'],
                    'nivel': nivel
                }
                carpetas_list.append(carpeta_info)
                # Recursión para subcarpetas
                listar_carpetas_recursivo(ctx, subfolder.properties['ServerRelativeUrl'], carpetas_list, nivel + 1)
    except Exception as e:
        st.warning(f"No se pudo acceder a {folder_url}: {str(e)}")

def obtener_archivos_sharepoint(ctx, folder_url):
    """Obtiene todos los archivos de una carpeta"""
    archivos = []
    try:
        folder = ctx.web.get_folder_by_server_relative_url(folder_url)
        files = folder.files
        ctx.load(files)
        ctx.execute_query()
        
        for file in files:
            # Filtrar solo documentos relevantes
            extension = Path(file.properties['Name']).suffix.lower()
            if extension in ['.pdf', '.docx', '.doc', '.txt', '.pptx']:
                archivos.append({
                    'nombre': file.properties['Name'],
                    'url': file.properties['ServerRelativeUrl'],
                    'tamaño': file.properties['Length'],
                    'modificado': file.properties['TimeLastModified']
                })
    except Exception as e:
        st.warning(f"Error leyendo archivos de {folder_url}: {str(e)}")
    
    return archivos

def descargar_archivo_sharepoint(ctx, file_url):
    """Descarga un archivo de SharePoint"""
    try:
        response = File.open_binary(ctx, file_url)
        return io.BytesIO(response.content)
    except Exception as e:
        st.error(f"Error descargando {file_url}: {str(e)}")
        return None

def procesar_archivo_sharepoint(ctx, archivo_info, embeddings):
    """Procesa un archivo individual de SharePoint"""
    try:
        # Descargar archivo
        file_content = descargar_archivo_sharepoint(ctx, archivo_info['url'])
        if not file_content:
            return None
        
        # Guardar temporalmente
        extension = Path(archivo_info['nombre']).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=extension) as tmp:
            tmp.write(file_content.read())
            tmp_path = tmp.name
        
        # Procesar según tipo
        if extension == '.pdf':
            from langchain_community.document_loaders import PyPDFLoader
            loader = PyPDFLoader(tmp_path)
        elif extension in ['.docx', '.doc']:
            from langchain_community.document_loaders import Docx2txtLoader
            loader = Docx2txtLoader(tmp_path)
        elif extension == '.txt':
            from langchain_community.document_loaders import TextLoader
            loader = TextLoader(tmp_path, encoding='utf-8')
        elif extension == '.pptx':
            from langchain_community.document_loaders import UnstructuredPowerPointLoader
            loader = UnstructuredPowerPointLoader(tmp_path)
        else:
            os.unlink(tmp_path)
            return None
        
        # Cargar documento
        docs = loader.load()
        
        # Añadir metadata
        for doc in docs:
            doc.metadata['fuente'] = archivo_info['nombre']
            doc.metadata['ruta_sharepoint'] = archivo_info['url']
        
        # Limpiar
        os.unlink(tmp_path)
        
        return docs
    
    except Exception as e:
        st.warning(f"Error procesando {archivo_info['nombre']}: {str(e)}")
        return None

def procesar_sharepoint_completo(ctx, carpetas_seleccionadas, embeddings):
    """Procesa todas las carpetas seleccionadas de SharePoint"""
    todos_documentos = []
    progreso = st.progress(0)
    status = st.empty()
    
    total_carpetas = len(carpetas_seleccionadas)
    
    for idx, carpeta in enumerate(carpetas_seleccionadas):
        status.text(f"📁 Procesando: {carpeta['nombre']}...")
        
        # Obtener archivos de la carpeta
        archivos = obtener_archivos_sharepoint(ctx, carpeta['url'])
        
        for archivo in archivos:
            status.text(f"📄 Procesando: {archivo['nombre']}...")
            docs = procesar_archivo_sharepoint(ctx, archivo, embeddings)
            if docs:
                todos_documentos.extend(docs)
                st.session_state.documentos_cargados.append(archivo['nombre'])
        
        progreso.progress((idx + 1) / total_carpetas)
    
    status.empty()
    progreso.empty()
    
    if not todos_documentos:
        st.error("No se encontraron documentos válidos")
        return None
    
    # Dividir en chunks
    splitter = RecursiveCharacterTextSplitter(
        chunk_size=1000,
        chunk_overlap=200,
        length_function=len
    )
    chunks = splitter.split_documents(todos_documentos)
    
    # Crear vectorstore
    vectorstore = FAISS.from_documents(chunks, embeddings)
    
    return vectorstore, len(chunks)

def crear_qa_chain(vectorstore, llm):
    """Crea cadena de Q&A"""
    template = """Eres un asistente empresarial que responde preguntas basándose en documentos de SharePoint.
Usa el contexto proporcionado para responder. Si no sabes la respuesta, dilo claramente.
Responde de manera profesional en español.

Contexto: {context}

Pregunta: {question}

Respuesta:"""
    
    prompt = PromptTemplate(template=template, input_variables=["context", "question"])
    retriever = vectorstore.as_retriever(search_kwargs={"k": 5})
    
    qa_chain = RetrievalQA.from_chain_type(
        llm=llm,
        chain_type="stuff",
        retriever=retriever,
        chain_type_kwargs={"prompt": prompt},
        return_source_documents=True
    )
    
    return qa_chain

# ===== INTERFAZ =====

st.title("📁 Sistema RAG - SharePoint")
st.markdown("Conecta a SharePoint y consulta tus documentos corporativos")

# Sidebar - Configuración SharePoint
with st.sidebar:
    st.header("🔐 Conexión SharePoint")
    
    with st.form("sharepoint_form"):
        site_url = st.text_input(
            "URL del sitio",
            placeholder="https://tuempresa.sharepoint.com/sites/tusite",
            help="URL completa del sitio SharePoint"
        )
        
        username = st.text_input(
            "Email corporativo",
            placeholder="usuario@tuempresa.com"
        )
        
        password = st.text_input(
            "Contraseña",
            type="password"
        )
        
        carpeta_raiz = st.text_input(
            "Carpeta raíz (opcional)",
            placeholder="/sites/tusite/Documentos compartidos",
            help="Deja vacío para usar 'Shared Documents'"
        )
        
        conectar_btn = st.form_submit_button("🔗 Conectar", type="primary")
    
    if conectar_btn and site_url and username and password:
        with st.spinner("Conectando a SharePoint..."):
            ctx = conectar_sharepoint(site_url, username, password)
            if ctx:
                st.session_state.sharepoint_ctx = ctx
                st.session_state.site_url = site_url
                st.session_state.carpeta_raiz = carpeta_raiz or "/Shared Documents"
                st.session_state.sharepoint_conectado = True
                st.success("✅ Conectado")
                st.rerun()
    
    st.divider()
    
    # Si está conectado, mostrar carpetas
    if st.session_state.sharepoint_conectado:
        st.success("✅ SharePoint conectado")
        
        if st.button("🔄 Actualizar carpetas"):
            st.session_state.carpetas_disponibles = None
        
        # Listar carpetas
        if 'carpetas_disponibles' not in st.session_state or st.session_state.carpetas_disponibles is None:
            with st.spinner("Cargando estructura de carpetas..."):
                carpetas = []
                listar_carpetas_recursivo(
                    st.session_state.sharepoint_ctx,
                    st.session_state.carpeta_raiz,
                    carpetas
                )
                st.session_state.carpetas_disponibles = carpetas
        
        st.subheader("📂 Seleccionar carpetas")
        
        if st.session_state.carpetas_disponibles:
            carpetas_seleccionadas = []
            
            for carpeta in st.session_state.carpetas_disponibles:
                indent = "  " * carpeta['nivel']
                if st.checkbox(
                    f"{indent}📁 {carpeta['nombre']}",
                    key=carpeta['url']
                ):
                    carpetas_seleccionadas.append(carpeta)
            
            st.divider()
            
            if carpetas_seleccionadas:
                st.info(f"✓ {len(carpetas_seleccionadas)} carpeta(s) seleccionada(s)")
                
                if st.button("🚀 Procesar documentos", type="primary"):
                    embeddings = inicializar_embeddings()
                    
                    with st.spinner("Procesando documentos de SharePoint..."):
                        resultado = procesar_sharepoint_completo(
                            st.session_state.sharepoint_ctx,
                            carpetas_seleccionadas,
                            embeddings
                        )
                        
                        if resultado:
                            vectorstore, num_chunks = resultado
                            st.session_state.vectorstore = vectorstore
                            
                            # Crear chain
                            llm = inicializar_llm()
                            st.session_state.qa_chain = crear_qa_chain(vectorstore, llm)
                            
                            st.success(f"✅ {len(st.session_state.documentos_cargados)} documentos procesados")
                            st.success(f"✅ {num_chunks} fragmentos indexados")
        else:
            st.warning("No se encontraron carpetas")
        
        if st.button("🔌 Desconectar"):
            st.session_state.sharepoint_conectado = False
            st.session_state.vectorstore = None
            st.session_state.qa_chain = None
            st.rerun()

# Área principal - Chat
if st.session_state.qa_chain:
    st.header("💬 Consulta tus documentos")
    
    # Info de documentos cargados
    with st.expander(f"📚 Documentos cargados ({len(st.session_state.documentos_cargados)})"):
        for doc in st.session_state.documentos_cargados:
            st.text(f"• {doc}")
    
    # Historial
    for pregunta, respuesta in st.session_state.historial:
        with st.chat_message("user"):
            st.write(pregunta)
        with st.chat_message("assistant"):
            st.write(respuesta)
    
    # Input
    pregunta = st.chat_input("Pregunta sobre tus documentos...")
    
    if pregunta:
        with st.chat_message("user"):
            st.write(pregunta)
        
        with st.chat_message("assistant"):
            with st.spinner("Buscando respuesta..."):
                resultado = st.session_state.qa_chain.invoke({"query": pregunta})
                respuesta = resultado['result']
                fuentes = resultado['source_documents']
                
                st.write(respuesta)
                
                with st.expander("📎 Fuentes"):
                    for i, doc in enumerate(fuentes, 1):
                        st.markdown(f"**{doc.metadata.get('fuente', 'Desconocido')}**")
                        st.text(doc.page_content[:200] + "...")
                        st.divider()
        
        st.session_state.historial.append((pregunta, respuesta))

else:
    st.info("👈 Conecta a SharePoint y selecciona carpetas para comenzar")
    
    st.markdown("""
    ### 📖 Cómo usar:
    
    1. **Conecta a SharePoint** con tus credenciales corporativas
    2. **Selecciona las carpetas** que quieres indexar
    3. **Procesa los documentos** (PDF, Word, PowerPoint, TXT)
    4. **Haz preguntas** y obtén respuestas basadas en tus documentos
    
    ### 🔒 Seguridad:
    - Las credenciales no se almacenan
    - Los documentos se procesan en memoria
    - Compatible con autenticación Microsoft 365
    """)

st.divider()
st.caption("Sistema RAG para SharePoint | Acceso corporativo seguro")
