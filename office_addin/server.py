"""
Office Add-in Server - Excel IA Chat
Servidor Flask con HTTPS que sirve el add-in y conecta con Claude/Gemini API
para generar código Office.js desde lenguaje natural.

Compatible con: Excel PC, Mac, iPad, Android y Web
"""

from flask import Flask, request, jsonify, send_from_directory
from flask_cors import CORS
import os
import re
import json
import ssl

from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))

app = Flask(__name__, static_folder='static')
CORS(app)

# ============================================================================
# CONFIGURACIÓN DE IA
# ============================================================================

# Prioridad: Claude > Gemini (configurable con ADDIN_IA_PROVIDER en .env)
IA_PROVIDER = os.getenv("ADDIN_IA_PROVIDER", "auto")  # "claude", "gemini", "auto"

# Claude
CLAUDE_API_KEY = os.getenv("ANTHROPIC_API_KEY", "")
CLAUDE_MODEL = os.getenv("CLAUDE_MODEL", "claude-sonnet-4-20250514")

# Gemini
GEMINI_API_KEY = os.getenv("GOOGLE_API_KEY", "") or os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "gemini-2.5-flash")


def get_ia_provider():
    """Determina qué proveedor de IA usar"""
    if IA_PROVIDER == "claude" and CLAUDE_API_KEY:
        return "claude"
    if IA_PROVIDER == "gemini" and GEMINI_API_KEY:
        return "gemini"
    if IA_PROVIDER == "auto":
        if CLAUDE_API_KEY:
            return "claude"
        if GEMINI_API_KEY:
            return "gemini"
    return None


# ============================================================================
# SYSTEM PROMPT PARA GENERAR OFFICE.JS
# ============================================================================

SYSTEM_PROMPT = """Eres un asistente experto en Microsoft Excel y Office.js (Office Add-ins JavaScript API).
El usuario te da instrucciones en lenguaje natural sobre una hoja de Excel y tú generas código JavaScript con Office.js para ejecutar la acción.

REGLAS ESTRICTAS:
1. Responde SIEMPRE en español
2. Genera código JavaScript válido usando la API de Office.js (Excel JavaScript API)
3. El código se ejecutará dentro de un Excel.run(), con las variables `context` y `sheet` ya disponibles
4. Envuelve el código en un bloque ```javascript ... ```
5. NO uses alert(), prompt(), confirm() ni formularios - solo modifica datos silenciosamente
6. Sé conciso en la explicación
7. NUNCA borres datos sin que el usuario lo pida explícitamente
8. Al final del código, SIEMPRE incluye `await context.sync();`
9. La variable `sheet` ya es la hoja activa (context.workbook.worksheets.getActiveWorksheet())
10. Cuando el usuario pida una operación que genera resultados nuevos (suma, multiplicación, fórmula, etc.) y NO especifique dónde ponerlos, SIEMPRE coloca el resultado en la SIGUIENTE COLUMNA VACÍA después de los datos existentes. Usa el contexto para saber cuántas columnas hay y pon los resultados en la columna siguiente.
11. SIEMPRE añade un encabezado descriptivo en la primera fila de la nueva columna (ej: "Suma A+B", "Total", etc.)
12. Analiza bien el CONTEXTO proporcionado (encabezados, rango, muestra de datos) para identificar correctamente qué columnas usar. Si el usuario dice "suma columna A con B", usa los datos reales de esas columnas.
13. PREFERIR SIEMPRE fórmulas de Excel (usando .formulas) en vez de valores calculados (usando .values). Por ejemplo, para sumar A+B, usa fórmulas como "=A2+B2" en vez de calcular el valor en JavaScript. Así los resultados se actualizan automáticamente si cambian los datos originales. Solo usa valores calculados cuando sea imposible hacerlo con fórmulas.
14. MUY IMPORTANTE - CARGA DE PROPIEDADES: En Office.js, ANTES de leer CUALQUIER propiedad de un objeto (como .values, .rowCount, .columnCount, .address, .text, .name, etc.), SIEMPRE debes:
    a) Llamar a .load() con las propiedades que necesitas: objeto.load(["propiedad1", "propiedad2"])
    b) Llamar a await context.sync() DESPUÉS del load y ANTES de leer los valores
    c) Solo DESPUÉS del context.sync() puedes acceder a los valores
    Ejemplo correcto:
      const range = sheet.getUsedRange();
      range.load(["values", "rowCount", "columnCount", "address"]);
      await context.sync();
      // Ahora sí puedes usar range.values, range.rowCount, etc.
    Ejemplo INCORRECTO (NUNCA hagas esto):
      const range = sheet.getUsedRange();
      const rows = range.rowCount; // ERROR: no se hizo load ni sync antes
15. Cuando necesites consultar datos (buscar un valor, calcular un promedio, encontrar el máximo, etc.), carga los valores primero con load+sync, haz el cálculo en JavaScript, y devuelve el resultado como texto en una celda o como respuesta.

CONTEXTO: El usuario te enviará información sobre:
- Nombre de la hoja activa
- Rango de datos utilizado (ej: A1:D10)
- Encabezados de columnas
- Muestra de datos (primeras filas)
- Número de filas y columnas

EJEMPLOS DE CÓDIGO VÁLIDO:

Ejemplo 1 - Sumar valor a una columna (columna B):
```javascript
const range = sheet.getRange("B2:B100");
range.load("values");
await context.sync();
const newValues = range.values.map(row => [row[0] !== "" && row[0] !== null ? row[0] + 10 : row[0]]);
range.values = newValues;
await context.sync();
```

Ejemplo 2 - Crear fórmula en nueva columna:
```javascript
const lastCol = sheet.getRange("1:1").getUsedRange();
lastCol.load("columnCount");
await context.sync();
const newCol = lastCol.columnCount;
const headerCell = sheet.getCell(0, newCol);
headerCell.values = [["Total"]];
const dataRange = sheet.getRangeByIndexes(1, newCol, 50, 1);
const formulas = [];
for (let i = 2; i <= 51; i++) {
    formulas.push(["=B" + i + "*C" + i]);
}
dataRange.formulas = formulas;
await context.sync();
```

Ejemplo 3 - Ordenar datos:
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load("address");
await context.sync();
const sortRange = usedRange;
sortRange.sort.apply([{ key: 0, ascending: true }]);
await context.sync();
```

Ejemplo 4 - Colorear filas que cumplen condición (Precio > 100 en columna B):
```javascript
const range = sheet.getUsedRange();
range.load(["values", "rowCount", "columnCount"]);
await context.sync();
for (let i = 1; i < range.rowCount; i++) {
    if (range.values[i][1] > 100) {
        const row = sheet.getRangeByIndexes(i, 0, 1, range.columnCount);
        row.format.fill.color = "#90EE90";
    }
}
await context.sync();
```

Ejemplo 5 - Filtrar (ocultar filas que no cumplen):
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load(["values", "rowCount"]);
await context.sync();
for (let i = 1; i < usedRange.rowCount; i++) {
    if (usedRange.values[i][1] <= 100) {
        const row = sheet.getRangeByIndexes(i, 0, 1, 1).getEntireRow();
        row.rowHidden = true;
    }
}
await context.sync();
```

Ejemplo 6 - Crear gráfico de barras con todos los datos:
```javascript
const dataRange = sheet.getUsedRange();
dataRange.load("address");
await context.sync();
const chart = sheet.charts.add("ColumnClustered", dataRange, "Auto");
chart.title.text = "Grafico de datos";
chart.legend.position = "Bottom";
chart.setPosition("G2", "N15");
await context.sync();
```

Ejemplo 7 - Crear gráfico de líneas con columnas específicas (A y B):
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load("rowCount");
await context.sync();
const chartDataRange = sheet.getRange("A1:B" + usedRange.rowCount);
const chart = sheet.charts.add("Line", chartDataRange, "Auto");
chart.title.text = "Grafico de lineas";
chart.legend.position = "Bottom";
chart.setPosition("G2", "N15");
await context.sync();
```

Ejemplo 8 - Crear gráfico circular:
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load("rowCount");
await context.sync();
const chartDataRange = sheet.getRange("A1:B" + usedRange.rowCount);
const chart = sheet.charts.add("Pie", chartDataRange, "Auto");
chart.title.text = "Grafico circular";
chart.legend.position = "Bottom";
chart.setPosition("G2", "N15");
await context.sync();
```

TIPOS DE GRAFICOS DISPONIBLES: "ColumnClustered", "ColumnStacked", "BarClustered", "BarStacked", "Line", "LineMarkers", "Pie", "Doughnut", "Area", "AreaStacked", "XYScatter", "XYScatterLines", "Radar"

IMPORTANTE SOBRE GRAFICOS:
- NUNCA uses métodos como getColumnLetter(), toColumn(), o funciones inventadas
- Para referenciar rangos usa sheet.getRange("A1:B10") con letras de columna directas
- Siempre haz load() y sync() antes de leer propiedades como rowCount
- Usa setPosition("G2", "N15") para posicionar el gráfico (primera celda, última celda)

Ejemplo 9 - Crear tabla dinámica (PivotTable):
```javascript
const usedRange = sheet.getUsedRange();
usedRange.load("address");
await context.sync();
const pivotSheet = context.workbook.worksheets.add("TablaDinamica");
const pivotTable = pivotSheet.pivotTables.add("MiPivot", usedRange, "A1");
pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Categoria"));
pivotTable.dataHierarchies.add(pivotTable.hierarchies.getItem("Ventas"));
await context.sync();
```

IMPORTANTE SOBRE TABLAS DINAMICAS:
- Usa context.workbook.worksheets.add() para crear una hoja nueva para la tabla dinámica
- Los nombres de los campos en getItem() deben coincidir EXACTAMENTE con los encabezados del contexto proporcionado
- Usa rowHierarchies.add() para filas, columnHierarchies.add() para columnas, y dataHierarchies.add() para valores
- Para funciones de agregación usa: pivotTable.dataHierarchies.getItem("Campo").summarizeBy = "Sum" (opciones: Sum, Count, Average, Max, Min)
"""


# ============================================================================
# FUNCIONES DE IA
# ============================================================================

def llamar_claude(instruccion, contexto_excel, historial):
    """Llama a la API de Claude (Anthropic)"""
    import anthropic
    client = anthropic.Anthropic(api_key=CLAUDE_API_KEY)

    messages = []

    if contexto_excel:
        messages.append({
            "role": "user",
            "content": f"CONTEXTO DE LA HOJA EXCEL ACTIVA:\n{contexto_excel}"
        })
        messages.append({
            "role": "assistant",
            "content": "Entendido. Tengo el contexto de tu hoja Excel. ¿Qué necesitas hacer?"
        })

    for msg in historial[-6:]:
        role = "user" if msg.get('rol') == 'usuario' else "assistant"
        messages.append({"role": role, "content": msg.get('texto', '')})

    messages.append({"role": "user", "content": instruccion})

    response = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=2048,
        system=SYSTEM_PROMPT,
        messages=messages
    )

    return response.content[0].text


def llamar_gemini(instruccion, contexto_excel, historial):
    """Llama a la API de Google Gemini"""
    from google import genai

    client = genai.Client(api_key=GEMINI_API_KEY)

    prompt_parts = [SYSTEM_PROMPT + "\n\n"]

    if contexto_excel:
        prompt_parts.append(f"CONTEXTO DE LA HOJA EXCEL ACTIVA:\n{contexto_excel}\n\n")

    for msg in historial[-6:]:
        prefix = "Usuario" if msg.get('rol') == 'usuario' else "Asistente"
        prompt_parts.append(f"{prefix}: {msg.get('texto', '')}\n")

    prompt_parts.append(f"Usuario: {instruccion}")

    response = client.models.generate_content(
        model=GEMINI_MODEL,
        contents="".join(prompt_parts)
    )

    return response.text


def extraer_codigo_js(texto):
    """Extrae código JavaScript de un bloque markdown"""
    match = re.search(r'```javascript\s*\n?(.*?)\n?\s*```', texto, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()

    match = re.search(r'```js\s*\n?(.*?)\n?\s*```', texto, re.DOTALL | re.IGNORECASE)
    if match:
        return match.group(1).strip()

    return None


# ============================================================================
# RUTAS - ARCHIVOS ESTÁTICOS DEL ADD-IN
# ============================================================================

@app.route('/')
def index():
    return send_from_directory('static', 'taskpane.html')


@app.route('/<path:filename>')
def static_files(filename):
    return send_from_directory('static', filename)


# ============================================================================
# RUTAS - API
# ============================================================================

@app.route('/api/health', methods=['GET'])
def health():
    provider = get_ia_provider()
    return jsonify({
        "status": "ok",
        "ia_provider": provider,
        "ia_disponible": provider is not None
    })


@app.route('/api/chat', methods=['POST'])
def chat():
    """Recibe instrucción + contexto de Excel y devuelve código Office.js"""
    try:
        data = request.get_json()
        instruccion = data.get('instruccion', '')
        contexto_excel = data.get('contexto', '')
        historial = data.get('historial', [])

        if not instruccion:
            return jsonify({"error": "No se proporcionó instrucción"}), 400

        provider = get_ia_provider()
        if not provider:
            return jsonify({
                "error": "No hay API de IA configurada. Añade ANTHROPIC_API_KEY o GOOGLE_API_KEY en el archivo .env"
            }), 500

        # Llamar al proveedor de IA
        if provider == "claude":
            respuesta_texto = llamar_claude(instruccion, contexto_excel, historial)
        else:
            respuesta_texto = llamar_gemini(instruccion, contexto_excel, historial)

        # Extraer código JavaScript
        codigo_js = extraer_codigo_js(respuesta_texto)

        return jsonify({
            "respuesta": respuesta_texto,
            "codigo_js": codigo_js,
            "tiene_codigo": codigo_js is not None,
            "provider": provider
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ============================================================================
# MAIN
# ============================================================================

if __name__ == '__main__':
    import sys
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

    provider = get_ia_provider()

    print("=" * 55)
    print("  Office Add-in Server - Excel IA Chat")
    print("  IA: %s" % (provider or "NO CONFIGURADA"))
    if provider == "claude":
        print("  Modelo: %s" % CLAUDE_MODEL)
    elif provider == "gemini":
        print("  Modelo: %s" % GEMINI_MODEL)
    print("  Puerto: 3000 (HTTPS)")
    print("=" * 55)

    if not provider:
        print()
        print("  No hay API de IA configurada!")
        print("  Anade una de estas claves en el archivo .env:")
        print("    ANTHROPIC_API_KEY=sk-ant-...")
        print("    GOOGLE_API_KEY=AI...")
        print()

    print()
    print("  Endpoints:")
    print("    GET  https://localhost:3000/         (Add-in UI)")
    print("    GET  https://localhost:3000/api/health")
    print("    POST https://localhost:3000/api/chat")
    print()
    print("  Presiona Ctrl+C para detener")
    print("=" * 55)

    # Verificar certificados SSL (priorizar los de office-addin-dev-certs)
    home = os.path.expanduser('~')
    ms_cert_dir = os.path.join(home, '.office-addin-dev-certs')
    local_cert_dir = os.path.join(os.path.dirname(__file__), 'certs')

    cert_file = None
    key_file = None

    # Primero buscar certificados oficiales de Microsoft
    ms_cert = os.path.join(ms_cert_dir, 'localhost.crt')
    ms_key = os.path.join(ms_cert_dir, 'localhost.key')
    if os.path.exists(ms_cert) and os.path.exists(ms_key):
        cert_file = ms_cert
        key_file = ms_key
        print("  Usando certificados de office-addin-dev-certs")
    else:
        # Buscar certificados locales
        local_cert = os.path.join(local_cert_dir, 'server.crt')
        local_key = os.path.join(local_cert_dir, 'server.key')
        if os.path.exists(local_cert) and os.path.exists(local_key):
            cert_file = local_cert
            key_file = local_key
            print("  Usando certificados locales")

    if cert_file and key_file:
        ssl_ctx = ssl.SSLContext(ssl.PROTOCOL_TLS_SERVER)
        ssl_ctx.load_cert_chain(cert_file, key_file)
        app.run(host='0.0.0.0', port=3000, debug=False, ssl_context=ssl_ctx)
    else:
        print()
        print("  No se encontraron certificados SSL.")
        print("  Ejecuta: npx office-addin-dev-certs install")
        print()
        print("  Iniciando en HTTP (solo para pruebas)...")
        print()
        app.run(host='0.0.0.0', port=3000, debug=False)
