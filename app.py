"""
Gestor de Excel con Streamlit + IA Híbrida
Aplicación para cargar, editar, analizar y exportar archivos Excel
con asistente de IA conversacional (Ollama local o Google Gemini)
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
import re
import os
import json

# Cargar variables de entorno
from dotenv import load_dotenv
load_dotenv(override=True)

# Compatibilidad con Streamlit Cloud secrets
def get_env(key, default=""):
    """Lee variable de entorno desde .env o st.secrets (Streamlit Cloud)"""
    val = os.getenv(key, "")
    if not val:
        try:
            val = st.secrets.get(key, default)
        except Exception:
            val = default
    return val

# ============================================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================================
st.set_page_config(
    page_title="📊 Gestor de Excel con IA",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# PWA: metadatos para añadir a pantalla de inicio sin barra de navegacion
st.markdown("""
<link rel="manifest" href="/app/static/manifest.json">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Excel IA">
<meta name="theme-color" content="#217346">
""", unsafe_allow_html=True)

# Estilos CSS personalizados
st.markdown("""
<style>
    .main { padding: 2rem; }
    .stDataFrame { border-radius: 10px; }
    .success-box {
        padding: 1rem;
        background-color: #d4edda;
        border-radius: 8px;
        border-left: 4px solid #28a745;
        margin: 0.5rem 0;
    }
    .info-box {
        padding: 1rem;
        background-color: #e3f2fd;
        border-radius: 8px;
        border-left: 4px solid #2196f3;
        margin: 0.5rem 0;
    }
    .chat-user {
        background-color: #e8f4f8;
        padding: 0.8rem;
        border-radius: 10px;
        margin: 0.5rem 0;
    }
    .chat-bot {
        background-color: #f0f2f6;
        padding: 0.8rem;
        border-radius: 10px;
        margin: 0.5rem 0;
    }
    .ia-badge {
        display: inline-block;
        padding: 0.2rem 0.5rem;
        border-radius: 4px;
        font-size: 0.8rem;
        font-weight: bold;
    }
    .ia-ollama { background-color: #e8f5e9; color: #2e7d32; }
    .ia-gemini { background-color: #e3f2fd; color: #1565c0; }
    .ia-claude { background-color: #fce4d6; color: #8b4513; }
    .ia-comandos { background-color: #fff3e0; color: #ef6c00; }

    /* Mejoras para iPad/móvil - Apple Pencil y táctil */
    @media (pointer: coarse), (hover: none) {
        /* Botones más grandes en móvil */
        .stButton button {
            min-height: 48px;
            font-size: 16px;
        }
        /* Inputs más grandes para evitar zoom automático en iOS */
        input, textarea, select {
            font-size: 16px !important;
        }
        /* Mejor área táctil para las celdas de la tabla */
        .stDataFrame td, .stDataFrame th {
            padding: 10px 8px !important;
            touch-action: manipulation;
        }
        /* Desactivar doble-tap zoom en tablas */
        .stDataFrame {
            touch-action: pan-x pan-y;
        }
    }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# INICIALIZACIÓN DEL ESTADO
# ============================================================================
if 'df' not in st.session_state:
    st.session_state.df = None
if 'filename' not in st.session_state:
    st.session_state.filename = None
if 'chat_history' not in st.session_state:
    st.session_state.chat_history = []
if 'historial_cambios' not in st.session_state:
    st.session_state.historial_cambios = []
if 'ia_provider' not in st.session_state:
    st.session_state.ia_provider = "Comandos"
if 'ollama_available' not in st.session_state:
    st.session_state.ollama_available = False
if 'gemini_available' not in st.session_state:
    st.session_state.gemini_available = False

# ============================================================================
# VERIFICAR DISPONIBILIDAD DE PROVEEDORES IA
# ============================================================================
@st.cache_resource
def check_ollama():
    """Verifica si Ollama está disponible"""
    try:
        import ollama
        modelos = ollama.list()
        # Verificar que hay al menos un modelo
        return len(modelos.get('models', [])) > 0 or hasattr(modelos, 'models')
    except Exception as e:
        print(f"Error verificando Ollama: {e}")
        return False

def check_gemini():
    """Verifica si Gemini está configurado"""
    api_key = get_env("GEMINI_API_KEY", "") or get_env("GOOGLE_API_KEY", "")
    return len(api_key) > 10 and api_key != "tu_api_key_aqui"

def check_claude():
    """Verifica si Claude está configurado"""
    api_key = get_env("ANTHROPIC_API_KEY", "")
    return len(api_key) > 10

st.session_state.ollama_available = check_ollama()
st.session_state.gemini_available = check_gemini()
if 'claude_available' not in st.session_state:
    st.session_state.claude_available = False
st.session_state.claude_available = check_claude()

# ============================================================================
# FUNCIONES AUXILIARES
# ============================================================================

def guardar_cambio(descripcion: str) -> None:
    """Guarda un registro del cambio realizado"""
    if st.session_state.df is not None:
        st.session_state.historial_cambios.append({
            'descripcion': descripcion,
            'df_copia': st.session_state.df.copy()
        })

def obtener_columnas_numericas() -> list:
    """Retorna lista de columnas numéricas del DataFrame"""
    if st.session_state.df is not None:
        return st.session_state.df.select_dtypes(include=['number']).columns.tolist()
    return []

def obtener_columnas_fecha() -> list:
    """Retorna lista de columnas de tipo fecha/datetime"""
    if st.session_state.df is not None:
        return st.session_state.df.select_dtypes(include=['datetime64', 'datetime64[ns]']).columns.tolist()
    return []

def obtener_todas_columnas() -> list:
    """Retorna lista de todas las columnas del DataFrame"""
    if st.session_state.df is not None:
        return st.session_state.df.columns.tolist()
    return []

def exportar_excel() -> BytesIO:
    """Exporta el DataFrame actual a un archivo Excel en memoria"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df.to_excel(writer, index=False, sheet_name='Datos')
    output.seek(0)
    return output

def obtener_contexto_df() -> str:
    """Genera contexto del DataFrame para la IA"""
    df = st.session_state.df
    if df is None:
        return "No hay datos cargados."
    
    cols_fecha = obtener_columnas_fecha()
    info_fechas = ""
    if cols_fecha:
        info_fechas = f"\n- Columnas de fecha: {cols_fecha} (puedes extraer año, mes, día o calcular diferencias)"
    
    info = f"""
INFORMACIÓN DEL DATAFRAME:
- Filas: {len(df)}
- Columnas: {list(df.columns)}
- Tipos de datos: {df.dtypes.to_dict()}{info_fechas}

PRIMERAS 5 FILAS:
{df.head().to_string()}

ESTADÍSTICAS NUMÉRICAS:
{df.describe().to_string() if len(obtener_columnas_numericas()) > 0 else "No hay columnas numéricas"}
"""
    return info

def crear_system_prompt() -> str:
    """Crea el prompt de sistema para la IA"""
    return """Eres un asistente experto en análisis de datos con pandas. El usuario tiene un DataFrame cargado y quiere manipularlo con lenguaje natural.

REGLAS:
1. Responde siempre en español
2. Si el usuario pide modificar datos, genera código Python usando la variable 'df'
3. Envuelve el código en un bloque ```python ... ```
4. El código debe ser seguro (no importar módulos, no acceder a archivos, no ejecutar comandos del sistema)
5. Si solo es una pregunta o consulta (media, suma, máximo, mínimo, contar, etc.), responde DIRECTAMENTE con el resultado calculado en tu respuesta de texto. NO generes código Python para consultas. Usa los datos del contexto proporcionado para calcular la respuesta.
6. Sé conciso pero amable
7. NUNCA uses print() en el código. El código solo debe modificar el DataFrame df.
8. Si los datos numéricos pueden estar como texto, convierte primero con pd.to_numeric(df['columna'], errors='coerce') antes de operar.
9. Los valores vacíos (NaN) se ignoran automáticamente en operaciones como mean(), sum(), etc.

EJEMPLOS DE CÓDIGO VÁLIDO:
- Sumar valor: df['Columna'] = df['Columna'] + 10
- Filtrar: df = df[df['Columna'] > 50]
- Ordenar: df = df.sort_values('Columna')
- Nueva columna: df['Nueva'] = df['A'] + df['B']
- Eliminar columna: df = df.drop(columns=['Columna'])
- Convertir a numérico: df['Columna'] = pd.to_numeric(df['Columna'], errors='coerce')

IMPORTANTE - OPERACIONES CON FECHAS:
- Las columnas de fecha pueden estar como texto (str). SIEMPRE usa pd.to_datetime(..., dayfirst=True, format='mixed') antes de operar con fechas.
- Diferencia en días: df['Diferencia'] = (pd.to_datetime(df['FechaA'], dayfirst=True, format='mixed') - pd.to_datetime(df['FechaB'], dayfirst=True, format='mixed')).dt.days
- Extraer año: df['Año'] = pd.to_datetime(df['Fecha'], dayfirst=True, format='mixed').dt.year
- Extraer mes: df['Mes'] = pd.to_datetime(df['Fecha'], dayfirst=True, format='mixed').dt.month

FUNCIONES DISPONIBLES:
- Para estadísticas: df.describe(), df.mean(), df.sum(), df.count()
- Para información: df.info(), df.shape, df.columns
- Para fechas: pd.to_datetime() (SIEMPRE usarlo antes de operaciones con fechas)"""

def extraer_codigo_python(texto: str) -> str | None:
    """Extrae código Python de un bloque markdown - soporta varios formatos"""
    # Primero limpiar tags de pensamiento de deepseek-r1
    texto_limpio = re.sub(r'<think>.*?</think>', '', texto, flags=re.DOTALL)
    
    # Intentar varios patrones de código
    patrones = [
        r'```python\s*\n?(.*?)\n?\s*```',  # ```python ... ```
        r'```py\s*\n?(.*?)\n?\s*```',      # ```py ... ```  
        r'```\s*\n?(df.*?)\n?\s*```',       # ``` df... ``` (sin especificar lenguaje)
    ]
    
    for patron in patrones:
        match = re.search(patron, texto_limpio, re.DOTALL | re.IGNORECASE)
        if match:
            return match.group(1).strip()
    
    # Buscar líneas que parecen código pandas fuera de bloques
    lineas_codigo = []
    for linea in texto_limpio.split('\n'):
        linea_limpia = linea.strip()
        # Detectar asignaciones de DataFrame
        if (linea_limpia.startswith('df[') or 
            linea_limpia.startswith('df =') or 
            linea_limpia.startswith('df.') or
            "df['" in linea_limpia or
            'df["' in linea_limpia):
            # Limpiar caracteres extra
            linea_limpia = linea_limpia.rstrip('.,;')
            if linea_limpia:
                lineas_codigo.append(linea_limpia)
    
    if lineas_codigo:
        return '\n'.join(lineas_codigo)
    
    return None

def ejecutar_codigo_seguro(codigo: str) -> tuple[bool, str]:
    """Ejecuta código Python de forma segura y retorna (éxito, mensaje)"""
    # Validar código peligroso
    palabras_prohibidas = ['import ', 'exec(', 'eval(', 'open(', 'os.', 'subprocess', 
                           'globals(', 'locals(', 'compile(', '__import__']
    
    for palabra in palabras_prohibidas:
        if palabra in codigo:
            return False, f"❌ Código no permitido: contiene '{palabra}'"
    
    try:
        # Guardar estado anterior para comparar
        cols_antes = len(st.session_state.df.columns)
        filas_antes = len(st.session_state.df)
        
        # Trabajar directamente con el DataFrame de session_state
        df = st.session_state.df
        
        # Builtins seguros necesarios para operaciones pandas
        safe_builtins = {
            'len': len,
            'str': str,
            'int': int,
            'float': float,
            'bool': bool,
            'list': list,
            'dict': dict,
            'range': range,
            'sum': sum,
            'min': min,
            'max': max,
            'abs': abs,
            'round': round,
            'sorted': sorted,
            'zip': zip,
            'enumerate': enumerate,
            'print': print,
            'True': True,
            'False': False,
            'None': None,
        }
        
        # Wrapper de pd que auto-configura to_datetime con dayfirst=True
        class PdWrapper:
            """Wrapper de pandas que mejora pd.to_datetime para fechas dd/mm/yyyy"""
            def __getattr__(self, name):
                return getattr(pd, name)
            def to_datetime(self, *args, **kwargs):
                # Intentar primero con la configuración proporcionada (o defaults)
                kwargs_intento = kwargs.copy()
                if 'dayfirst' not in kwargs_intento:
                    kwargs_intento['dayfirst'] = True
                if 'format' not in kwargs_intento:
                    kwargs_intento['format'] = 'mixed'
                
                try:
                    return pd.to_datetime(*args, **kwargs_intento)
                except Exception:
                    # Si falla (por ejemplo, formato incorrecto), reintentar forzando mixed
                    kwargs_fallback = kwargs.copy()
                    kwargs_fallback['dayfirst'] = True
                    kwargs_fallback['format'] = 'mixed' 
                    return pd.to_datetime(*args, **kwargs_fallback)
        
        # El df aquí es una referencia directa, no una copia
        local_vars = {'df': df, 'pd': PdWrapper()}
        
        exec(codigo, {"__builtins__": safe_builtins}, local_vars)
        
        # Si el código reasignó df (ej: df = df.drop(...)), actualizar session_state
        if local_vars['df'] is not df:
            st.session_state.df = local_vars['df']
        
        # Verificar cambios
        cols_despues = len(st.session_state.df.columns)
        filas_despues = len(st.session_state.df)
        
        cambios = []
        if cols_despues != cols_antes:
            cambios.append(f"columnas: {cols_antes} → {cols_despues}")
        if filas_despues != filas_antes:
            cambios.append(f"filas: {filas_antes} → {filas_despues}")
        
        guardar_cambio(f"Código IA ejecutado")
        
        if cambios:
            return True, f"✅ Ejecutado. Cambios: {', '.join(cambios)}"
        else:
            return True, "✅ Ejecutado (valores modificados)"
        
    except Exception as e:
        return False, f"❌ Error: {str(e)}\n\nCódigo intentado:\n```\n{codigo}\n```"

# ============================================================================
# FUNCIONES DE IA
# ============================================================================

def chat_ollama(mensaje: str, historial: list) -> str:
    """Envía mensaje a Ollama y retorna respuesta"""
    try:
        import ollama
        
        modelo = get_env("OLLAMA_MODEL", "qwen2.5:3b")
        contexto = obtener_contexto_df()
        
        messages = [
            {"role": "system", "content": crear_system_prompt()},
            {"role": "user", "content": f"CONTEXTO DEL DATAFRAME:\n{contexto}"}
        ]
        
        # Agregar historial reciente (últimos 6 mensajes)
        for msg in historial[-6:]:
            role = "user" if msg['rol'] == 'usuario' else "assistant"
            messages.append({"role": role, "content": msg['texto']})
        
        # Agregar mensaje actual
        messages.append({"role": "user", "content": mensaje})
        
        response = ollama.chat(model=modelo, messages=messages)
        return response['message']['content']
        
    except Exception as e:
        return f"❌ Error con Ollama: {str(e)}"

def chat_gemini(mensaje: str, historial: list) -> str:
    """Envía mensaje a Google Gemini y retorna respuesta"""
    try:
        from google import genai

        api_key = get_env("GEMINI_API_KEY") or get_env("GOOGLE_API_KEY")
        client = genai.Client(api_key=api_key)

        modelo = get_env("GEMINI_MODEL", "gemini-2.5-flash")
        contexto = obtener_contexto_df()
        system_prompt = crear_system_prompt()

        # Construir prompt completo
        prompt_completo = f"""{system_prompt}

CONTEXTO DEL DATAFRAME:
{contexto}

HISTORIAL DE CONVERSACIÓN:
"""
        for msg in historial[-6:]:
            rol = "Usuario" if msg['rol'] == 'usuario' else "Asistente"
            prompt_completo += f"{rol}: {msg['texto']}\n"

        prompt_completo += f"\nUsuario: {mensaje}\nAsistente:"

        response = client.models.generate_content(
            model=modelo,
            contents=prompt_completo
        )

        return response.text

    except Exception as e:
        return f"❌ Error con Gemini: {str(e)}"

def chat_claude(mensaje: str, historial: list) -> str:
    """Envía mensaje a Claude (Anthropic) y retorna respuesta"""
    try:
        import anthropic

        api_key = get_env("ANTHROPIC_API_KEY")
        client = anthropic.Anthropic(api_key=api_key)

        modelo = get_env("CLAUDE_MODEL", "claude-sonnet-4-20250514")
        contexto = obtener_contexto_df()
        system_prompt = crear_system_prompt()

        messages = []

        if contexto:
            messages.append({
                "role": "user",
                "content": f"CONTEXTO DEL DATAFRAME:\n{contexto}"
            })
            messages.append({
                "role": "assistant",
                "content": "Entendido. Tengo el contexto de tus datos. ¿Qué necesitas hacer?"
            })

        for msg in historial[-6:]:
            role = "user" if msg['rol'] == 'usuario' else "assistant"
            messages.append({"role": role, "content": msg['texto']})

        messages.append({"role": "user", "content": mensaje})

        response = client.messages.create(
            model=modelo,
            max_tokens=2048,
            system=system_prompt,
            messages=messages
        )

        return response.content[0].text

    except Exception as e:
        return f"❌ Error con Claude: {str(e)}"

def procesar_comando_regex(comando: str) -> str:
    """Procesa comandos con regex (modo sin IA)"""
    comando = comando.lower().strip()
    df = st.session_state.df
    
    if df is None:
        return "❌ Primero debes cargar un archivo Excel."
    
    columnas = obtener_todas_columnas()
    columnas_lower = [c.lower() for c in columnas]
    
    # Comando: SUMA
    patron_suma = r'suma\s+(\d+(?:\.\d+)?)\s+(?:a\s+)?(.+)'
    match = re.search(patron_suma, comando)
    if match:
        numero = float(match.group(1))
        col_buscar = match.group(2).strip()
        
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                if df[col_real].dtype in ['int64', 'float64']:
                    guardar_cambio(f"Suma {numero} a '{col_real}'")
                    st.session_state.df[col_real] = df[col_real] + numero
                    return f"✅ Se sumó {numero} a la columna '{col_real}'."
                else:
                    return f"❌ La columna '{col_real}' no es numérica."
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: RESTA
    patron_resta = r'resta\s+(\d+(?:\.\d+)?)\s+(?:a\s+)?(.+)'
    match = re.search(patron_resta, comando)
    if match:
        numero = float(match.group(1))
        col_buscar = match.group(2).strip()
        
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                if df[col_real].dtype in ['int64', 'float64']:
                    guardar_cambio(f"Resta {numero} a '{col_real}'")
                    st.session_state.df[col_real] = df[col_real] - numero
                    return f"✅ Se restó {numero} a la columna '{col_real}'."
                else:
                    return f"❌ La columna '{col_real}' no es numérica."
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: MULTIPLICA
    patron_mult = r'multiplic[a|ar]\s+(.+?)\s+por\s+(\d+(?:\.\d+)?)'
    match = re.search(patron_mult, comando)
    if match:
        col_buscar = match.group(1).strip()
        numero = float(match.group(2))
        
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                if df[col_real].dtype in ['int64', 'float64']:
                    guardar_cambio(f"Multiplica '{col_real}' por {numero}")
                    st.session_state.df[col_real] = df[col_real] * numero
                    return f"✅ Se multiplicó la columna '{col_real}' por {numero}."
                else:
                    return f"❌ La columna '{col_real}' no es numérica."
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: FILTRA
    patron_filtra = r'filtra\s+(?:donde\s+)?(.+?)\s*([><=!]+)\s*(\d+(?:\.\d+)?)'
    match = re.search(patron_filtra, comando)
    if match:
        col_buscar = match.group(1).strip()
        operador = match.group(2)
        valor = float(match.group(3))
        
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                guardar_cambio(f"Filtrar '{col_real}' {operador} {valor}")
                
                if operador == '>':
                    st.session_state.df = df[df[col_real] > valor]
                elif operador == '<':
                    st.session_state.df = df[df[col_real] < valor]
                elif operador == '>=':
                    st.session_state.df = df[df[col_real] >= valor]
                elif operador == '<=':
                    st.session_state.df = df[df[col_real] <= valor]
                elif operador == '==' or operador == '=':
                    st.session_state.df = df[df[col_real] == valor]
                
                filas = len(st.session_state.df)
                return f"✅ Filtro aplicado. Quedan {filas} filas."
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: ORDENA
    patron_ordena = r'ordena\s+(?:por\s+)?(.+?)(?:\s+(asc|desc|ascendente|descendente))?$'
    match = re.search(patron_ordena, comando)
    if match:
        col_buscar = match.group(1).strip()
        orden = match.group(2) if match.group(2) else 'asc'
        ascendente = 'desc' not in orden.lower()
        
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                guardar_cambio(f"Ordenar por '{col_real}'")
                st.session_state.df = st.session_state.df.sort_values(
                    by=col_real, ascending=ascendente
                ).reset_index(drop=True)
                direccion = "ascendente" if ascendente else "descendente"
                return f"✅ Tabla ordenada por '{col_real}' ({direccion})."
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: ESTADÍSTICAS
    if 'estadística' in comando or 'resumen' in comando or 'describe' in comando:
        cols_num = obtener_columnas_numericas()
        if cols_num:
            stats = df[cols_num].describe()
            return f"📊 Estadísticas:\n```\n{stats.to_string()}\n```"
        return "❌ No hay columnas numéricas."
    
    # Comando: EXTRAER AÑO
    patron_year = r'extra[e|er]\s+(?:el\s+)?año\s+(?:de\s+)?(.+)'
    match = re.search(patron_year, comando)
    if match:
        col_buscar = match.group(1).strip()
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                try:
                    st.session_state.df[f'Año_{col_real}'] = pd.to_datetime(df[col_real]).dt.year
                    guardar_cambio(f"Extraer año de '{col_real}'")
                    return f"✅ Columna 'Año_{col_real}' creada con el año extraído."
                except:
                    return f"❌ No pude extraer el año de '{col_real}'. ¿Es una fecha válida?"
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: EXTRAER MES
    patron_month = r'extra[e|er]\s+(?:el\s+)?mes\s+(?:de\s+)?(.+)'
    match = re.search(patron_month, comando)
    if match:
        col_buscar = match.group(1).strip()
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                try:
                    st.session_state.df[f'Mes_{col_real}'] = pd.to_datetime(df[col_real]).dt.month
                    guardar_cambio(f"Extraer mes de '{col_real}'")
                    return f"✅ Columna 'Mes_{col_real}' creada con el mes extraído."
                except:
                    return f"❌ No pude extraer el mes de '{col_real}'. ¿Es una fecha válida?"
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: EXTRAER DÍA
    patron_day = r'extra[e|er]\s+(?:el\s+)?día\s+(?:de\s+)?(.+)'
    match = re.search(patron_day, comando)
    if match:
        col_buscar = match.group(1).strip()
        for i, col in enumerate(columnas_lower):
            if col_buscar in col or col in col_buscar:
                col_real = columnas[i]
                try:
                    st.session_state.df[f'Día_{col_real}'] = pd.to_datetime(df[col_real]).dt.day
                    guardar_cambio(f"Extraer día de '{col_real}'")
                    return f"✅ Columna 'Día_{col_real}' creada con el día extraído."
                except:
                    return f"❌ No pude extraer el día de '{col_real}'. ¿Es una fecha válida?"
        return f"❌ No encontré la columna '{col_buscar}'."
    
    # Comando: AYUDA
    if 'ayuda' in comando or 'help' in comando:
        return """📚 **Comandos disponibles:**

**Operaciones numéricas:**
• `suma 10 a Precio`
• `resta 5 a Cantidad`
• `multiplica Ventas por 2`
• `filtra Precio > 50`
• `ordena por Nombre`
• `estadísticas`

**Operaciones con fechas:** 📅
• `extrae año de Fecha`
• `extrae mes de Fecha`
• `extrae día de Fecha`"""
    
    return "🤔 No entendí. Escribe 'ayuda' para ver comandos."

def procesar_mensaje_ia(mensaje: str) -> tuple[str, bool]:
    """Procesa mensaje con el proveedor de IA seleccionado. Retorna (respuesta, hubo_cambios)"""
    provider = st.session_state.ia_provider
    historial = st.session_state.chat_history
    
    if provider == "Comandos":
        respuesta = procesar_comando_regex(mensaje)
        hubo_cambios = "✅" in respuesta
        return respuesta, hubo_cambios
    
    elif provider == "Ollama":
        respuesta = chat_ollama(mensaje, historial)
        
        # Verificar si hay código para ejecutar
        codigo = extraer_codigo_python(respuesta)
        
        if codigo:
            # Mostrar qué código se detectó
            debug_msg = f"\n\n---\n🔍 **Código detectado:**\n```python\n{codigo}\n```\n"
            exito, msg_ejecucion = ejecutar_codigo_seguro(codigo)
            return f"{respuesta}{debug_msg}\n{msg_ejecucion}", exito
        else:
            return f"{respuesta}\n\n---\n⚠️ No se detectó código ejecutable en la respuesta.", False
    
    elif provider == "Gemini":
        respuesta = chat_gemini(mensaje, historial)

        codigo = extraer_codigo_python(respuesta)
        if codigo:
            debug_msg = f"\n\n---\n🔍 **Código detectado:**\n```python\n{codigo}\n```\n"
            exito, msg_ejecucion = ejecutar_codigo_seguro(codigo)
            return f"{respuesta}{debug_msg}\n{msg_ejecucion}", exito
        else:
            return f"{respuesta}\n\n---\n⚠️ No se detectó código ejecutable.", False

    elif provider == "Claude":
        respuesta = chat_claude(mensaje, historial)

        codigo = extraer_codigo_python(respuesta)
        if codigo:
            debug_msg = f"\n\n---\n🔍 **Código detectado:**\n```python\n{codigo}\n```\n"
            exito, msg_ejecucion = ejecutar_codigo_seguro(codigo)
            return f"{respuesta}{debug_msg}\n{msg_ejecucion}", exito
        else:
            return f"{respuesta}\n\n---\n⚠️ No se detectó código ejecutable.", False

    return "❌ Proveedor de IA no válido", False

# ============================================================================
# BARRA LATERAL
# ============================================================================
with st.sidebar:
    st.header("📂 Cargar Archivo")

    archivo = st.file_uploader(
        "Selecciona un archivo Excel",
        # Sin filtro 'type' para compatibilidad con Android/pCloud (el nombre puede no tener extensión)
        help="Formatos: .xlsx, .xls — Desde pCloud u otras nubes en Android, abre primero la app pCloud y 'Compartir' el archivo"
    )

    # Solo cargar si es un archivo NUEVO (diferente al ya cargado)
    if archivo is not None:
        nombre_lower = archivo.name.lower()

        # MIME types válidos para Excel (pCloud y otras apps pueden omitir la extensión)
        MIME_EXCEL = {
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',  # xlsx
            'application/vnd.ms-excel',                                            # xls
            'application/msexcel',
            'application/x-msexcel',
            'application/octet-stream',  # genérico (pCloud a veces lo usa)
        }

        # Validar: extensión correcta O tipo MIME reconocido como Excel
        tiene_extension_excel = nombre_lower.endswith('.xlsx') or nombre_lower.endswith('.xls')
        tiene_mime_excel = archivo.type in MIME_EXCEL

        es_excel_valido = tiene_extension_excel or tiene_mime_excel

        if not es_excel_valido:
            st.error(f"❌ No se reconoce como Excel.")
            st.caption(f"📄 Nombre: `{archivo.name}`")
            st.caption(f"📎 Tipo detectado: `{archivo.type}`")
            st.caption("💡 Si viene de pCloud, intenta descargarlo primero a la tablet y luego súbelo desde ahí.")
        else:
            # Verificar si es un archivo diferente al actual
            archivo_nuevo = st.session_state.filename != archivo.name

            if archivo_nuevo:
                try:
                    df_nuevo = pd.read_excel(archivo)
                    # Auto-convertir columnas que parecen numéricas
                    for col in df_nuevo.columns:
                        if df_nuevo[col].dtype == object or str(df_nuevo[col].dtype) == 'string':
                            try:
                                converted = df_nuevo[col].astype(str).str.replace(',', '.', regex=False)
                                converted = pd.to_numeric(converted, errors='coerce')
                                non_null = df_nuevo[col].dropna()
                                if len(non_null) > 0:
                                    num_valid = converted.dropna().count()
                                    if num_valid / len(non_null) >= 0.5:
                                        df_nuevo[col] = converted
                            except Exception:
                                pass
                    st.session_state.df = df_nuevo
                    st.session_state.filename = archivo.name
                    st.session_state.chat_history = []
                    st.success(f"✅ Cargado: {archivo.name}")
                except Exception as e:
                    st.error(f"❌ Error al leer el archivo: {str(e)}")
                    st.caption(f"📄 Nombre: `{archivo.name}` | Tipo: `{archivo.type}`")
                    st.caption("💡 Tip: Si viene de pCloud, descarga el archivo a la tablet primero.")
    
    # Selector de proveedor IA
    st.divider()
    st.header("🤖 Proveedor de IA")
    
    opciones_ia = ["Comandos"]
    if st.session_state.ollama_available:
        opciones_ia.append("Ollama")
    if st.session_state.gemini_available:
        opciones_ia.append("Gemini")
    if st.session_state.claude_available:
        opciones_ia.append("Claude")

    st.session_state.ia_provider = st.radio(
        "Selecciona el asistente:",
        opciones_ia,
        help="Ollama = local, Gemini = cloud, Claude = cloud, Comandos = regex"
    )

    # Estado de proveedores
    with st.expander("ℹ️ Estado de proveedores"):
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.session_state.ollama_available:
                st.success("🟢 Ollama")
            else:
                st.warning("🔴 Ollama")
        with col2:
            if st.session_state.gemini_available:
                st.success("🟢 Gemini")
            else:
                st.warning("🔴 Gemini")
        with col3:
            if st.session_state.claude_available:
                st.success("🟢 Claude")
            else:
                st.warning("🔴 Claude")

        if not st.session_state.gemini_available:
            st.caption("Configura GEMINI_API_KEY en .env")
        if not st.session_state.claude_available:
            st.caption("Configura ANTHROPIC_API_KEY en .env")
    
    # Info del archivo
    if st.session_state.df is not None:
        st.divider()
        st.subheader("📋 Información")
        df = st.session_state.df
        
        col1, col2 = st.columns(2)
        col1.metric("Filas", len(df))
        col2.metric("Columnas", len(df.columns))
        
        with st.expander("Ver columnas"):
            for col in df.columns:
                tipo = str(df[col].dtype)
                if "datetime" in tipo:
                    icono = "📅"
                elif "int" in tipo or "float" in tipo:
                    icono = "🔢"
                else:
                    icono = "📝"
                st.write(f"{icono} **{col}** ({tipo})")
        
        # Botón de descarga
        st.divider()
        excel_data = exportar_excel()
        nombre_desc = st.session_state.filename.replace('.xlsx', '_modificado.xlsx')
        st.download_button(
            label="📥 Descargar Excel",
            data=excel_data,
            file_name=nombre_desc,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ============================================================================
# CONTENIDO PRINCIPAL
# ============================================================================
st.title("📊 Gestor de Excel con IA")

# Badge del proveedor actual
provider = st.session_state.ia_provider
if provider == "Ollama":
    st.markdown("<span class='ia-badge ia-ollama'>🦙 Ollama (Local)</span>", unsafe_allow_html=True)
elif provider == "Gemini":
    st.markdown("<span class='ia-badge ia-gemini'>✨ Gemini (Cloud)</span>", unsafe_allow_html=True)
elif provider == "Claude":
    st.markdown("<span class='ia-badge ia-claude'>🤖 Claude (Cloud)</span>", unsafe_allow_html=True)
else:
    st.markdown("<span class='ia-badge ia-comandos'>⌨️ Comandos</span>", unsafe_allow_html=True)

if st.session_state.df is None:
    st.info("👈 Carga un archivo Excel desde la barra lateral para comenzar.")
    st.stop()

# Pestañas principales
tab1, tab2, tab3, tab4, tab5 = st.tabs(["📝 Editar Datos", "🧮 Fórmulas", "📈 Gráficos", "🔄 Tabla Dinámica", "🤖 Chat IA"])

# ============================================================================
# PESTAÑA 1: EDITOR DE DATOS
# ============================================================================
with tab1:
    st.subheader("Editor de Datos")

    # Toggle para modo edición (evita teclado en móvil)
    if 'edit_mode' not in st.session_state:
        st.session_state.edit_mode = False

    st.session_state.edit_mode = st.toggle("Activar edición", value=st.session_state.edit_mode,
                                            help="Desactiva para ver datos sin que salte el teclado en iPad/móvil")

    if st.session_state.edit_mode:
        st.caption("Modo edición: toca una celda para editarla")
        df_editado = st.data_editor(
            st.session_state.df,
            use_container_width=True,
            num_rows="dynamic",
            height=500
        )
        if not df_editado.equals(st.session_state.df):
            st.session_state.df = df_editado
            st.toast("✏️ Datos actualizados", icon="✅")
    else:
        st.caption("Modo lectura: activa la edición arriba para modificar celdas")
        st.dataframe(
            st.session_state.df,
            use_container_width=True,
            height=500
        )

# ============================================================================
# PESTAÑA 2: FÓRMULAS
# ============================================================================
with tab2:
    st.subheader("Aplicar Fórmulas")

    # Layout dividido: controles (izquierda) | datos en tiempo real (derecha)
    col_form_ctrl, col_form_data = st.columns([3, 2])

    with col_form_ctrl:
        cols_numericas = obtener_columnas_numericas()

        if not cols_numericas:
            st.warning("⚠️ No hay columnas numéricas.")
        else:
            col1, col2, col3 = st.columns([1, 1, 1])

            with col1:
                operacion = st.selectbox("Operación", ["Suma", "Resta", "Multiplicación", "División", "Promedio"])
            with col2:
                columna_a = st.selectbox("Columna A", cols_numericas, key="col_a")
            with col3:
                columna_b = st.selectbox("Columna B", cols_numericas, key="col_b")

            nombre_resultado = st.text_input("Nombre columna resultado", value=f"Resultado_{operacion}")

            if st.button("➕ Aplicar Fórmula", type="primary"):
                df = st.session_state.df
                guardar_cambio(f"Fórmula: {columna_a} {operacion} {columna_b}")

                if operacion == "Suma":
                    df[nombre_resultado] = df[columna_a] + df[columna_b]
                elif operacion == "Resta":
                    df[nombre_resultado] = df[columna_a] - df[columna_b]
                elif operacion == "Multiplicación":
                    df[nombre_resultado] = df[columna_a] * df[columna_b]
                elif operacion == "División":
                    df[nombre_resultado] = df[columna_a] / df[columna_b].replace(0, float('nan'))
                elif operacion == "Promedio":
                    df[nombre_resultado] = (df[columna_a] + df[columna_b]) / 2

                st.session_state.df = df
                st.success(f"✅ Columna '{nombre_resultado}' creada.")
                st.rerun()

        # Sección de operaciones con fechas
        st.divider()
        st.subheader("📅 Operaciones con Fechas")

        cols_fecha = obtener_columnas_fecha()

        if not cols_fecha:
            st.info("ℹ️ No hay columnas de fecha detectadas. Las fechas deben estar en formato datetime.")
            st.caption("💡 Tip: Si tienes fechas como texto, la IA puede ayudarte a convertirlas.")
        else:
            operacion_fecha = st.selectbox(
                "Operación con fechas",
                ["Extraer Año", "Extraer Mes", "Extraer Día", "Extraer Día de la Semana",
                 "Días entre dos fechas", "Añadir días", "Restar días"],
                key="op_fecha"
            )

            col_fecha1, col_fecha2 = st.columns(2)

            with col_fecha1:
                columna_fecha = st.selectbox("Columna de fecha", cols_fecha, key="col_fecha_1")

            with col_fecha2:
                if operacion_fecha == "Días entre dos fechas":
                    columna_fecha_2 = st.selectbox("Segunda fecha", cols_fecha, key="col_fecha_2")
                elif operacion_fecha in ["Añadir días", "Restar días"]:
                    cantidad_dias = st.number_input("Cantidad de días", min_value=1, value=7, key="cant_dias")

            nombre_resultado_fecha = st.text_input(
                "Nombre nueva columna",
                value=f"{operacion_fecha.replace(' ', '_')}_{columna_fecha}",
                key="nombre_col_fecha"
            )

            if st.button("📅 Aplicar operación", type="primary", key="btn_fecha"):
                df = st.session_state.df

                try:
                    if operacion_fecha == "Extraer Año":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]).dt.year
                        guardar_cambio(f"Extraer año de '{columna_fecha}'")

                    elif operacion_fecha == "Extraer Mes":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]).dt.month
                        guardar_cambio(f"Extraer mes de '{columna_fecha}'")

                    elif operacion_fecha == "Extraer Día":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]).dt.day
                        guardar_cambio(f"Extraer día de '{columna_fecha}'")

                    elif operacion_fecha == "Extraer Día de la Semana":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]).dt.day_name()
                        guardar_cambio(f"Extraer día semana de '{columna_fecha}'")

                    elif operacion_fecha == "Días entre dos fechas":
                        df[nombre_resultado_fecha] = (pd.to_datetime(df[columna_fecha_2]) - pd.to_datetime(df[columna_fecha])).dt.days
                        guardar_cambio(f"Días entre '{columna_fecha}' y '{columna_fecha_2}'")

                    elif operacion_fecha == "Añadir días":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]) + pd.Timedelta(days=cantidad_dias)
                        guardar_cambio(f"Añadir {cantidad_dias} días a '{columna_fecha}'")

                    elif operacion_fecha == "Restar días":
                        df[nombre_resultado_fecha] = pd.to_datetime(df[columna_fecha]) - pd.Timedelta(days=cantidad_dias)
                        guardar_cambio(f"Restar {cantidad_dias} días a '{columna_fecha}'")

                    st.session_state.df = df
                    st.success(f"✅ Columna '{nombre_resultado_fecha}' creada.")
                    st.rerun()

                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")

    # Columna derecha: datos en tiempo real
    with col_form_data:
        st.markdown("#### 📊 Datos en Tiempo Real")
        m1, m2 = st.columns(2)
        m1.metric("Filas", len(st.session_state.df))
        m2.metric("Columnas", len(st.session_state.df.columns))
        st.dataframe(st.session_state.df, use_container_width=True, height=480)
        nombre_archivo_f = st.session_state.filename.replace('.xlsx', '_modificado.xlsx') if st.session_state.filename else 'datos_modificados.xlsx'
        st.download_button(
            "📥 Exportar Excel Modificado",
            data=exportar_excel(),
            file_name=nombre_archivo_f,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_form_{len(st.session_state.df.columns)}_{len(st.session_state.df)}",
            use_container_width=True
        )

# ============================================================================
# PESTAÑA 3: GRÁFICOS AVANZADOS
# ============================================================================
with tab3:
    st.subheader("📊 Generador de Gráficos Avanzado")

    # Layout dividido: controles (izquierda) | datos en tiempo real (derecha)
    col_graf_ctrl, col_graf_data = st.columns([3, 2])

    with col_graf_ctrl:
        todas_columnas = obtener_todas_columnas()
        cols_numericas = obtener_columnas_numericas()
        cols_fecha = obtener_columnas_fecha()

        # Categorías de gráficos
        TIPOS_GRAFICOS = {
            "📊 Básicos": ["Barras", "Barras Horizontales", "Líneas", "Área", "Dispersión"],
            "📈 Comparación": ["Barras Agrupadas", "Barras Apiladas", "Área Apilada"],
            "🥧 Distribución": ["Pastel", "Anillo (Donut)", "Histograma", "Boxplot"],
            "🔥 Avanzados": ["Heatmap (Correlación)", "Treemap", "Radar", "Embudo", "Cascada"],
            "📅 Temporal": ["Serie Temporal"] if cols_fecha else []
        }

        # Aplanar opciones
        opciones_tipo = []
        for categoria, tipos in TIPOS_GRAFICOS.items():
            if tipos:
                opciones_tipo.extend(tipos)

        # Layout de controles
        col_tipo, col_config = st.columns([1, 2])

        with col_tipo:
            tipo_grafico = st.selectbox("Tipo de gráfico", opciones_tipo, key="tipo_graf")

            # Paletas de colores
            paletas = {
                "Plotly": px.colors.qualitative.Plotly,
                "Vibrante": px.colors.qualitative.Vivid,
                "Pastel": px.colors.qualitative.Pastel,
                "Oscuro": px.colors.qualitative.Dark24,
                "Set1": px.colors.qualitative.Set1,
                "Safe": px.colors.qualitative.Safe,
            }
            paleta_nombre = st.selectbox("Paleta de colores", list(paletas.keys()))
            paleta = paletas[paleta_nombre]

        with col_config:
            # Configuración según tipo de gráfico
            if tipo_grafico in ["Heatmap (Correlación)"]:
                st.info("📌 El Heatmap muestra correlaciones entre columnas numéricas.")
            elif tipo_grafico in ["Histograma", "Boxplot"]:
                eje_x = st.selectbox("Columna", cols_numericas if cols_numericas else todas_columnas)
                eje_y = None
            elif tipo_grafico in ["Pastel", "Anillo (Donut)", "Treemap", "Embudo"]:
                col_cat, col_val = st.columns(2)
                with col_cat:
                    eje_x = st.selectbox("Categoría", todas_columnas)
                with col_val:
                    eje_y = st.selectbox("Valor (opcional)", ["Conteo automático"] + cols_numericas)
            elif tipo_grafico == "Radar":
                eje_x = st.selectbox("Categorías", [c for c in todas_columnas if c not in cols_numericas])
                eje_y = st.multiselect("Variables numéricas", cols_numericas, default=cols_numericas[:3] if len(cols_numericas) >= 3 else cols_numericas)
            elif tipo_grafico == "Cascada":
                col_cat, col_val = st.columns(2)
                with col_cat:
                    eje_x = st.selectbox("Etiquetas", todas_columnas)
                with col_val:
                    eje_y = st.selectbox("Valores", cols_numericas)
            elif tipo_grafico in ["Barras Agrupadas", "Barras Apiladas", "Área Apilada"]:
                col1, col2, col3 = st.columns(3)
                with col1:
                    eje_x = st.selectbox("Eje X", todas_columnas)
                with col2:
                    eje_y = st.multiselect("Valores Y (múltiples)", cols_numericas, default=[cols_numericas[0]] if cols_numericas else [])
                with col3:
                    if tipo_grafico != "Área Apilada":
                        color_col = st.selectbox("Agrupar por (color)", ["Ninguno"] + [c for c in todas_columnas if c != eje_x])
                    else:
                        color_col = "Ninguno"
            elif tipo_grafico == "Serie Temporal":
                col1, col2 = st.columns(2)
                with col1:
                    eje_x = st.selectbox("Columna de Fecha", cols_fecha)
                with col2:
                    eje_y = st.multiselect("Valores a graficar", cols_numericas, default=[cols_numericas[0]] if cols_numericas else [])
            else:
                col1, col2 = st.columns(2)
                with col1:
                    eje_x = st.selectbox("Eje X", todas_columnas)
                with col2:
                    eje_y = st.selectbox("Eje Y", cols_numericas if cols_numericas else todas_columnas)

        # Opciones de personalización
        with st.expander("🎨 Personalización", expanded=False):
            col_p1, col_p2, col_p3 = st.columns(3)
            with col_p1:
                titulo_grafico = st.text_input("Título del gráfico", value="")
            with col_p2:
                mostrar_leyenda = st.checkbox("Mostrar leyenda", value=True)
            with col_p3:
                altura_grafico = st.slider("Altura (px)", 300, 800, 500)

        # Generar gráfico
        if st.button("📊 Generar Gráfico", type="primary", key="btn_graf_avanzado"):
            df = st.session_state.df.copy()

            # Título automático si no se especifica
            if not titulo_grafico:
                if tipo_grafico == "Heatmap (Correlación)":
                    titulo_grafico = "Mapa de Correlaciones"
                elif isinstance(eje_y, list):
                    titulo_grafico = f"{tipo_grafico}: {', '.join(eje_y)} por {eje_x}"
                elif eje_y:
                    titulo_grafico = f"{tipo_grafico}: {eje_y} por {eje_x}"
                else:
                    titulo_grafico = f"{tipo_grafico}: {eje_x}"

            try:
                fig = None

                # ===== GRÁFICOS BÁSICOS =====
                if tipo_grafico == "Barras":
                    fig = px.bar(df, x=eje_x, y=eje_y, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Barras Horizontales":
                    fig = px.bar(df, x=eje_y, y=eje_x, orientation='h', title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Líneas":
                    fig = px.line(df, x=eje_x, y=eje_y, title=titulo_grafico, color_discrete_sequence=paleta, markers=True)

                elif tipo_grafico == "Área":
                    fig = px.area(df, x=eje_x, y=eje_y, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Dispersión":
                    fig = px.scatter(df, x=eje_x, y=eje_y, title=titulo_grafico, color_discrete_sequence=paleta)

                # ===== GRÁFICOS DE COMPARACIÓN =====
                elif tipo_grafico == "Barras Agrupadas":
                    if color_col != "Ninguno":
                        fig = px.bar(df, x=eje_x, y=eje_y[0] if len(eje_y) == 1 else eje_y, color=color_col,
                                    barmode='group', title=titulo_grafico, color_discrete_sequence=paleta)
                    else:
                        df_melt = df.melt(id_vars=[eje_x], value_vars=eje_y, var_name='Serie', value_name='Valor')
                        fig = px.bar(df_melt, x=eje_x, y='Valor', color='Serie', barmode='group',
                                    title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Barras Apiladas":
                    if color_col != "Ninguno":
                        fig = px.bar(df, x=eje_x, y=eje_y[0] if len(eje_y) == 1 else eje_y, color=color_col,
                                    barmode='stack', title=titulo_grafico, color_discrete_sequence=paleta)
                    else:
                        df_melt = df.melt(id_vars=[eje_x], value_vars=eje_y, var_name='Serie', value_name='Valor')
                        fig = px.bar(df_melt, x=eje_x, y='Valor', color='Serie', barmode='stack',
                                    title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Área Apilada":
                    df_melt = df.melt(id_vars=[eje_x], value_vars=eje_y, var_name='Serie', value_name='Valor')
                    fig = px.area(df_melt, x=eje_x, y='Valor', color='Serie', title=titulo_grafico, color_discrete_sequence=paleta)

                # ===== GRÁFICOS DE DISTRIBUCIÓN =====
                elif tipo_grafico == "Pastel":
                    if eje_y == "Conteo automático":
                        conteo = df[eje_x].value_counts().reset_index()
                        conteo.columns = [eje_x, 'Cantidad']
                        fig = px.pie(conteo, names=eje_x, values='Cantidad', title=titulo_grafico, color_discrete_sequence=paleta)
                    else:
                        fig = px.pie(df, names=eje_x, values=eje_y, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Anillo (Donut)":
                    if eje_y == "Conteo automático":
                        conteo = df[eje_x].value_counts().reset_index()
                        conteo.columns = [eje_x, 'Cantidad']
                        fig = px.pie(conteo, names=eje_x, values='Cantidad', title=titulo_grafico, hole=0.4, color_discrete_sequence=paleta)
                    else:
                        fig = px.pie(df, names=eje_x, values=eje_y, title=titulo_grafico, hole=0.4, color_discrete_sequence=paleta)

                elif tipo_grafico == "Histograma":
                    fig = px.histogram(df, x=eje_x, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Boxplot":
                    fig = px.box(df, y=eje_x, title=titulo_grafico, color_discrete_sequence=paleta)

                # ===== GRÁFICOS AVANZADOS =====
                elif tipo_grafico == "Heatmap (Correlación)":
                    corr_matrix = df[cols_numericas].corr()
                    fig = px.imshow(corr_matrix, text_auto=True, title=titulo_grafico,
                                   color_continuous_scale='RdBu_r', aspect='auto')

                elif tipo_grafico == "Treemap":
                    if eje_y == "Conteo automático":
                        conteo = df[eje_x].value_counts().reset_index()
                        conteo.columns = [eje_x, 'Cantidad']
                        fig = px.treemap(conteo, path=[eje_x], values='Cantidad', title=titulo_grafico, color_discrete_sequence=paleta)
                    else:
                        fig = px.treemap(df, path=[eje_x], values=eje_y, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Radar":
                    if eje_y:
                        df_radar = df.groupby(eje_x)[eje_y].mean().reset_index()
                        fig = go.Figure()
                        for idx, row in df_radar.iterrows():
                            fig.add_trace(go.Scatterpolar(
                                r=[row[col] for col in eje_y],
                                theta=eje_y,
                                fill='toself',
                                name=str(row[eje_x])
                            ))
                        fig.update_layout(title=titulo_grafico, polar=dict(radialaxis=dict(visible=True)))

                elif tipo_grafico == "Embudo":
                    if eje_y == "Conteo automático":
                        conteo = df[eje_x].value_counts().reset_index()
                        conteo.columns = [eje_x, 'Cantidad']
                        fig = px.funnel(conteo, x='Cantidad', y=eje_x, title=titulo_grafico, color_discrete_sequence=paleta)
                    else:
                        fig = px.funnel(df, x=eje_y, y=eje_x, title=titulo_grafico, color_discrete_sequence=paleta)

                elif tipo_grafico == "Cascada":
                    fig = go.Figure(go.Waterfall(
                        x=df[eje_x].tolist(),
                        y=df[eje_y].tolist(),
                        connector={"line": {"color": "rgb(63, 63, 63)"}},
                    ))
                    fig.update_layout(title=titulo_grafico)

                # ===== GRÁFICOS TEMPORALES =====
                elif tipo_grafico == "Serie Temporal":
                    df = df.sort_values(by=eje_x)
                    if eje_y:
                        df_melt = df.melt(id_vars=[eje_x], value_vars=eje_y, var_name='Serie', value_name='Valor')
                        fig = px.line(df_melt, x=eje_x, y='Valor', color='Serie', title=titulo_grafico,
                                     markers=True, color_discrete_sequence=paleta)
                        fig.update_xaxes(tickformat="%d/%m/%Y")

                # Aplicar estilos comunes
                if fig:
                    fig.update_layout(
                        template="plotly_white",
                        font=dict(size=14),
                        title_font=dict(size=20),
                        height=altura_grafico,
                        showlegend=mostrar_leyenda
                    )

                    # Mostrar gráfico
                    st.plotly_chart(fig, use_container_width=True, key="grafico_principal")

                    # Botón de exportación
                    col_exp1, col_exp2, col_exp3 = st.columns([1, 1, 2])
                    with col_exp1:
                        import io
                        buffer = io.StringIO()
                        fig.write_html(buffer)
                        html_bytes = buffer.getvalue().encode()
                        st.download_button(
                            "📥 Descargar HTML",
                            data=html_bytes,
                            file_name=f"grafico_{tipo_grafico.replace(' ', '_')}.html",
                            mime="text/html"
                        )
                    with col_exp2:
                        st.caption("💡 El HTML es interactivo como el gráfico original")
                else:
                    st.warning("⚠️ No se pudo generar el gráfico con la configuración actual.")

            except Exception as e:
                st.error(f"❌ Error al generar gráfico: {str(e)}")

    # Columna derecha: datos en tiempo real
    with col_graf_data:
        st.markdown("#### 📊 Datos en Tiempo Real")
        m1, m2 = st.columns(2)
        m1.metric("Filas", len(st.session_state.df))
        m2.metric("Columnas", len(st.session_state.df.columns))
        st.dataframe(st.session_state.df, use_container_width=True, height=480)
        nombre_archivo_g = st.session_state.filename.replace('.xlsx', '_modificado.xlsx') if st.session_state.filename else 'datos_modificados.xlsx'
        st.download_button(
            "📥 Exportar Excel Modificado",
            data=exportar_excel(),
            file_name=nombre_archivo_g,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_graf_{len(st.session_state.df.columns)}_{len(st.session_state.df)}",
            use_container_width=True
        )

# ============================================================================
# PESTAÑA 4: TABLA DINÁMICA
# ============================================================================
with tab4:
    st.subheader("🔄 Tabla Dinámica (Pivot Table)")
    st.caption("Agrupa y resume tus datos como en Excel")

    # Layout dividido: controles (izquierda) | datos en tiempo real (derecha)
    col_piv_ctrl, col_piv_data = st.columns([3, 2])

    with col_piv_ctrl:
        todas_columnas = obtener_todas_columnas()
        cols_numericas = obtener_columnas_numericas()
        cols_categoricas = [c for c in todas_columnas if c not in cols_numericas]

        # Configuración de la tabla dinámica
        col_config1, col_config2 = st.columns(2)


        with col_config1:
            st.markdown("**📋 Estructura**")

            # Filas (índice)
            filas_pivot = st.multiselect(
                "Agrupar por (Filas)",
                todas_columnas,
                default=[cols_categoricas[0]] if cols_categoricas else [],
                help="Categorías para las filas de la tabla"
            )

            # Columnas (opcional)
            columnas_pivot = st.selectbox(
                "Separar por (Columnas) - Opcional",
                ["Ninguno"] + todas_columnas,
                help="Segunda categoría para crear columnas"
            )
            if columnas_pivot == "Ninguno":
                columnas_pivot = None

        with col_config2:
            st.markdown("**📊 Valores**")

            # Valores a agregar
            valores_pivot = st.multiselect(
                "Columnas de valores",
                cols_numericas if cols_numericas else todas_columnas,
                default=[cols_numericas[0]] if cols_numericas else [],
                help="Columnas numéricas a agregar"
            )

            # Función de agregación
            funciones_agg = {
                "Suma": "sum",
                "Promedio": "mean",
                "Conteo": "count",
                "Mínimo": "min",
                "Máximo": "max",
                "Desviación Estándar": "std",
                "Mediana": "median"
            }
            funcion_agg = st.selectbox("Función de agregación", list(funciones_agg.keys()))

        # Opciones adicionales
        with st.expander("⚙️ Opciones adicionales", expanded=False):
            col_opt1, col_opt2, col_opt3 = st.columns(3)
            with col_opt1:
                mostrar_totales = st.checkbox("Mostrar totales", value=True)
            with col_opt2:
                rellenar_nulos = st.checkbox("Rellenar vacíos con 0", value=True)
            with col_opt3:
                formato_numeros = st.selectbox("Formato números", ["Normal", "2 decimales", "Enteros", "Porcentaje"])

        # Generar tabla dinámica
        if st.button("🔄 Generar Tabla Dinámica", type="primary", key="btn_pivot"):
            if not filas_pivot:
                st.warning("⚠️ Selecciona al menos una columna para agrupar (Filas).")
            elif not valores_pivot:
                st.warning("⚠️ Selecciona al menos una columna de valores.")
            else:
                try:
                    df = st.session_state.df.copy()

                    # Crear tabla dinámica
                    tabla_pivot = pd.pivot_table(
                        df,
                        values=valores_pivot,
                        index=filas_pivot,
                        columns=columnas_pivot,
                        aggfunc=funciones_agg[funcion_agg],
                        margins=mostrar_totales,
                        margins_name="TOTAL",
                        fill_value=0 if rellenar_nulos else None
                    )

                    # Formatear números
                    if formato_numeros == "2 decimales":
                        tabla_formateada = tabla_pivot.round(2)
                    elif formato_numeros == "Enteros":
                        tabla_formateada = tabla_pivot.round(0).astype(int)
                    elif formato_numeros == "Porcentaje":
                        tabla_formateada = tabla_pivot.round(4) * 100
                    else:
                        tabla_formateada = tabla_pivot

                    # Mostrar resultado
                    st.success(f"✅ Tabla generada: {len(tabla_formateada)} filas")

                    # Mostrar tabla con estilo
                    st.dataframe(
                        tabla_formateada,
                        use_container_width=True,
                        height=400
                    )

                    # Botones de exportación
                    col_exp1, col_exp2, col_exp3 = st.columns([1, 1, 2])
                    with col_exp1:
                        # Exportar como Excel
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            tabla_formateada.to_excel(writer, sheet_name='Tabla Dinámica')
                        output.seek(0)
                        st.download_button(
                            "📥 Descargar Excel",
                            data=output,
                            file_name="tabla_dinamica.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    with col_exp2:
                        # Exportar como CSV
                        csv_data = tabla_formateada.to_csv().encode('utf-8')
                        st.download_button(
                            "📥 Descargar CSV",
                            data=csv_data,
                            file_name="tabla_dinamica.csv",
                            mime="text/csv"
                        )

                    # Gráfico de la tabla dinámica
                    st.divider()
                    st.markdown("**📊 Visualización**")

                    try:
                        if columnas_pivot:
                            fig = px.bar(
                                tabla_formateada.reset_index(),
                                x=filas_pivot[0] if len(filas_pivot) == 1 else tabla_formateada.reset_index().columns[0],
                                y=tabla_formateada.columns.tolist()[:10],
                                barmode='group',
                                title=f"{funcion_agg} de {', '.join(valores_pivot)}"
                            )
                        else:
                            fig = px.bar(
                                tabla_formateada.reset_index(),
                                x=filas_pivot[0] if len(filas_pivot) == 1 else tabla_formateada.reset_index().columns[0],
                                y=valores_pivot[0] if len(valores_pivot) == 1 else valores_pivot,
                                title=f"{funcion_agg} de {', '.join(valores_pivot)}"
                            )

                        fig.update_layout(template="plotly_white", height=400)
                        st.plotly_chart(fig, use_container_width=True)
                    except:
                        st.caption("💡 No se pudo generar el gráfico automático.")

                except Exception as e:
                    st.error(f"❌ Error al generar tabla: {str(e)}")

    # Columna derecha: datos en tiempo real
    with col_piv_data:
        st.markdown("#### 📊 Datos en Tiempo Real")
        m1, m2 = st.columns(2)
        m1.metric("Filas", len(st.session_state.df))
        m2.metric("Columnas", len(st.session_state.df.columns))
        st.dataframe(st.session_state.df, use_container_width=True, height=480)
        nombre_archivo_p = st.session_state.filename.replace('.xlsx', '_modificado.xlsx') if st.session_state.filename else 'datos_modificados.xlsx'
        st.download_button(
            "📥 Exportar Excel Modificado",
            data=exportar_excel(),
            file_name=nombre_archivo_p,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"dl_piv_{len(st.session_state.df.columns)}_{len(st.session_state.df)}",
            use_container_width=True
        )

# ============================================================================
# PESTAÑA 5: CHAT IA (Vista dividida: Chat + Datos en tiempo real)
# ============================================================================
with tab5:
    provider = st.session_state.ia_provider
    
    if provider == "Ollama":
        st.subheader("🦙 Chat con Ollama (Local)")
        st.caption("Habla naturalmente con la IA para manipular tus datos — los cambios se reflejan en tiempo real a la derecha")
    elif provider == "Gemini":
        st.subheader("✨ Chat con Gemini (Cloud)")
        st.caption("Habla naturalmente con la IA para manipular tus datos — los cambios se reflejan en tiempo real a la derecha")
    elif provider == "Claude":
        st.subheader("🤖 Chat con Claude (Cloud)")
        st.caption("Habla naturalmente con la IA para manipular tus datos — los cambios se reflejan en tiempo real a la derecha")
    else:
        st.subheader("⌨️ Modo Comandos")
        st.caption("Usa comandos estructurados (suma, resta, filtra, ordena...) — los cambios se reflejan en tiempo real a la derecha")
    
    # ==================== INPUT DE CHAT (ARRIBA) ====================
    placeholder = "Ejemplo: ¿Cuál es el promedio de ventas?" if provider != "Comandos" else "Ejemplo: suma 10 a Precio"
    comando = st.chat_input(placeholder)

    if comando:
        st.session_state.chat_history.append({'rol': 'usuario', 'texto': comando})

        with st.spinner("Pensando..."):
            respuesta, hubo_cambios = procesar_mensaje_ia(comando)

        st.session_state.chat_history.append({'rol': 'asistente', 'texto': respuesta})
        st.rerun()

    # Botones de acción
    col_btn1, col_btn2, col_btn3 = st.columns([1, 1, 2])
    with col_btn1:
        if st.session_state.chat_history:
            if st.button("🗑️ Limpiar chat"):
                st.session_state.chat_history = []
                st.rerun()
    with col_btn2:
        if st.session_state.historial_cambios:
            if st.button("↩️ Deshacer último"):
                if len(st.session_state.historial_cambios) > 0:
                    ultimo = st.session_state.historial_cambios.pop()
                    st.session_state.df = ultimo['df_copia']
                    st.toast(f"↩️ Deshecho: {ultimo['descripcion']}", icon="↩️")
                    st.rerun()

    # ---- Layout dividido: Chat (izquierda) | Datos en tiempo real (derecha) ----
    col_chat, col_datos = st.columns([3, 2])

    # ==================== COLUMNA DERECHA: DATOS EN TIEMPO REAL ====================
    with col_datos:
        st.markdown("#### 📊 Datos en Tiempo Real")

        m1, m2 = st.columns(2)
        m1.metric("Filas", len(st.session_state.df))
        m2.metric("Columnas", len(st.session_state.df.columns))

        st.dataframe(
            st.session_state.df,
            use_container_width=True,
            height=420
        )

        with st.expander("🔍 Detalle de columnas", expanded=False):
            for col in st.session_state.df.columns:
                tipo = str(st.session_state.df[col].dtype)
                if "datetime" in tipo:
                    icono = "📅"
                elif "int" in tipo or "float" in tipo:
                    icono = "🔢"
                else:
                    icono = "📝"
                st.caption(f"{icono} **{col}** ({tipo})")

        if st.session_state.historial_cambios:
            with st.expander(f"📝 Últimos cambios ({len(st.session_state.historial_cambios)})", expanded=False):
                for cambio in st.session_state.historial_cambios[-5:]:
                    st.caption(f"• {cambio['descripcion']}")

        nombre_archivo = st.session_state.filename.replace('.xlsx', '_modificado.xlsx') if st.session_state.filename else 'datos_modificados.xlsx'
        st.download_button(
            "📥 Exportar Excel Modificado",
            data=exportar_excel(),
            file_name=nombre_archivo,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"download_{len(st.session_state.df.columns)}_{len(st.session_state.df)}",
            use_container_width=True
        )

    # ==================== COLUMNA IZQUIERDA: HISTORIAL CHAT ====================
    with col_chat:
        chat_container = st.container(height=500)

        with chat_container:
            if not st.session_state.chat_history:
                st.info("💬 Escribe un mensaje arriba para empezar a interactuar con tus datos.")

            for mensaje in st.session_state.chat_history:
                if mensaje['rol'] == 'usuario':
                    st.markdown(f"<div class='chat-user'>👤 **Tú:** {mensaje['texto']}</div>",
                               unsafe_allow_html=True)
                else:
                    st.markdown(f"<div class='chat-bot'>🤖 **Asistente:** {mensaje['texto']}</div>",
                               unsafe_allow_html=True)
        
        # Ejemplos
        with st.expander("💡 Ejemplos de preguntas" if provider != "Comandos" else "💡 Comandos disponibles"):
            if provider != "Comandos":
                st.markdown("""
| Pregunta | Acción |
|----------|--------|
| ¿Cuál es el promedio de Precio? | Estadística |
| Suma 10 a la columna Ventas | Modificar |
| Filtra donde Precio > 100 | Filtrar |
| Crea columna Total = Precio * Cantidad | Nueva columna |
| Ordena por Fecha | Ordenar |
""")
            else:
                st.markdown("""
| Comando | Descripción |
|---------|-------------|
| `suma 10 a Precio` | Suma a columna |
| `resta 5 a Cantidad` | Resta a columna |
| `multiplica Ventas por 2` | Multiplica columna |
| `filtra Precio > 50` | Filtra filas |
| `ordena por Nombre` | Ordena tabla |
| `estadísticas` | Muestra resumen |
""")

# ============================================================================
# PIE DE PÁGINA
# ============================================================================
st.divider()
st.caption("📊 Gestor de Excel con IA | Ollama + Gemini + Claude + Streamlit")
