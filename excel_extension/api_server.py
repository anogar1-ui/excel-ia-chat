"""
API Server para Excel + Ollama
Servidor Flask que conecta Microsoft Excel (VBA) con Ollama para
manipular hojas de cálculo usando lenguaje natural.
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import ollama
import os
import re
import json

# Cargar variables de entorno
try:
    from dotenv import load_dotenv
    load_dotenv(os.path.join(os.path.dirname(__file__), '..', '.env'))
except ImportError:
    pass

app = Flask(__name__)
CORS(app)

OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen2.5-coder:7b")

SYSTEM_PROMPT = """Eres un asistente experto en Microsoft Excel y VBA. El usuario te da instrucciones en lenguaje natural sobre una hoja de Excel y tú generas código VBA para ejecutar la acción.

REGLAS ESTRICTAS:
1. Responde SIEMPRE en español
2. Genera código VBA válido y seguro
3. El código debe trabajar con la hoja activa (ActiveSheet)
4. Envuelve el código VBA en un bloque ```vba ... ```
5. NO uses InputBox, MsgBox ni formularios - solo modifica datos silenciosamente
6. Sé conciso en la explicación
7. NUNCA borres datos sin que el usuario lo pida explícitamente

CONTEXTO: El usuario te enviará información sobre:
- Nombre de la hoja activa
- Rango de datos utilizado (ej: A1:D10)
- Encabezados de columnas
- Muestra de datos (primeras filas)

EJEMPLOS DE CÓDIGO VBA VÁLIDO:

Ejemplo 1 - Sumar valor a una columna:
```vba
Dim ws As Worksheet
Set ws = ActiveSheet
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
Dim i As Long
For i = 2 To lastRow
    ws.Cells(i, 2).Value = ws.Cells(i, 2).Value + 10
Next i
```

Ejemplo 2 - Crear fórmula en nueva columna:
```vba
Dim ws As Worksheet
Set ws = ActiveSheet
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
ws.Cells(1, 5).Value = "Total"
Dim i As Long
For i = 2 To lastRow
    ws.Cells(i, 5).Formula = "=B" & i & "*C" & i
Next i
```

Ejemplo 3 - Ordenar datos:
```vba
Dim ws As Worksheet
Set ws = ActiveSheet
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Dim lastCol As Long
lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).Sort _
    Key1:=ws.Range("A1"), Order1:=xlAscending, Header:=xlYes
```

Ejemplo 4 - Filtrar (colorear filas que cumplen condición):
```vba
Dim ws As Worksheet
Set ws = ActiveSheet
Dim lastRow As Long
lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
Dim i As Long
For i = 2 To lastRow
    If ws.Cells(i, 2).Value > 100 Then
        ws.Range(ws.Cells(i, 1), ws.Cells(i, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).Interior.Color = RGB(144, 238, 144)
    End If
Next i
```
"""


def extraer_codigo_vba(texto):
    """Extrae código VBA de un bloque markdown y quita Sub/End Sub wrapper"""
    codigo = None
    
    # Buscar bloque ```vba ... ```
    match = re.search(r'```vba\s*\n?(.*?)\n?\s*```', texto, re.DOTALL | re.IGNORECASE)
    if match:
        codigo = match.group(1).strip()
    
    if not codigo:
        # Buscar bloque ``` ... ``` con contenido VBA
        match = re.search(r'```\s*\n?((?:(?:Dim|Set|Sub|For|If|ws\.).*?\n?)+.*?)```', texto, re.DOTALL)
        if match:
            codigo = match.group(1).strip()
    
    if not codigo:
        return None
    
    # Quitar Sub ... / End Sub wrapper (el VBA del cliente añade su propio wrapper)
    lineas = codigo.split('\n')
    lineas_limpias = []
    for linea in lineas:
        linea_strip = linea.strip()
        # Saltar líneas Sub XXX() y End Sub
        if re.match(r'^(Public\s+|Private\s+)?Sub\s+\w+\s*\(', linea_strip, re.IGNORECASE):
            continue
        if linea_strip.lower() == 'end sub':
            continue
        lineas_limpias.append(linea)
    
    resultado = '\n'.join(lineas_limpias).strip()
    return resultado if resultado else None


@app.route('/health', methods=['GET'])
def health():
    """Verificar que el servidor está activo"""
    return jsonify({"status": "ok", "model": OLLAMA_MODEL})


@app.route('/status', methods=['GET'])
def status():
    """Estado de Ollama y modelo"""
    try:
        modelos = ollama.list()
        nombres = [m.get('name', m.get('model', '')) for m in modelos.get('models', [])]
        modelo_disponible = any(OLLAMA_MODEL in n for n in nombres)
        return jsonify({
            "ollama": True,
            "model": OLLAMA_MODEL,
            "model_loaded": modelo_disponible,
            "available_models": nombres
        })
    except Exception as e:
        return jsonify({
            "ollama": False,
            "error": str(e)
        }), 500


@app.route('/chat', methods=['POST'])
def chat():
    """Recibe instrucción + contexto de Excel y devuelve código VBA"""
    try:
        data = request.get_json()
        
        instruccion = data.get('instruccion', '')
        contexto_excel = data.get('contexto', '')
        historial = data.get('historial', [])
        
        if not instruccion:
            return jsonify({"error": "No se proporcionó instrucción"}), 400
        
        # Construir mensajes para Ollama
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT}
        ]
        
        # Agregar contexto del Excel
        if contexto_excel:
            messages.append({
                "role": "user", 
                "content": f"CONTEXTO DE LA HOJA EXCEL ACTIVA:\n{contexto_excel}"
            })
        
        # Agregar historial reciente
        for msg in historial[-4:]:
            role = "user" if msg.get('rol') == 'usuario' else "assistant"
            messages.append({"role": role, "content": msg.get('texto', '')})
        
        # Agregar instrucción actual
        messages.append({"role": "user", "content": instruccion})
        
        # Llamar a Ollama
        response = ollama.chat(model=OLLAMA_MODEL, messages=messages)
        respuesta_texto = response['message']['content']
        
        # Extraer código VBA
        codigo_vba = extraer_codigo_vba(respuesta_texto)
        
        return jsonify({
            "respuesta": respuesta_texto,
            "codigo_vba": codigo_vba,
            "tiene_codigo": codigo_vba is not None
        })
        
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route('/chat_simple', methods=['POST'])
def chat_simple():
    """Versión simplificada que devuelve texto plano (para VBA)"""
    try:
        data = request.get_json()
        
        instruccion = data.get('instruccion', '')
        contexto_excel = data.get('contexto', '')
        
        if not instruccion:
            return "===TEXTO===\nError: No se proporcionó instrucción\n===FIN===", 200, {'Content-Type': 'text/plain; charset=utf-8'}
        
        # Construir mensajes para Ollama
        messages = [
            {"role": "system", "content": SYSTEM_PROMPT},
        ]
        
        if contexto_excel:
            messages.append({
                "role": "user", 
                "content": f"CONTEXTO DE LA HOJA EXCEL ACTIVA:\n{contexto_excel}"
            })
        
        messages.append({"role": "user", "content": instruccion})
        
        # Llamar a Ollama
        response = ollama.chat(model=OLLAMA_MODEL, messages=messages)
        respuesta_texto = response['message']['content']
        
        # Extraer código VBA
        codigo_vba = extraer_codigo_vba(respuesta_texto)
        
        # Devolver texto plano con delimitadores simples
        resultado = "===TEXTO===\n"
        resultado += respuesta_texto + "\n"
        if codigo_vba:
            resultado += "===CODIGO===\n"
            resultado += codigo_vba + "\n"
        resultado += "===FIN==="
        
        return resultado, 200, {'Content-Type': 'text/plain; charset=utf-8'}
        
    except Exception as e:
        return f"===TEXTO===\nError: {str(e)}\n===FIN===", 200, {'Content-Type': 'text/plain; charset=utf-8'}


if __name__ == '__main__':
    print("=" * 50)
    print("  🤖 Excel + Ollama API Server")
    print(f"  Modelo: {OLLAMA_MODEL}")
    print("  Puerto: 5050")
    print("=" * 50)
    print()
    print("  El servidor está listo para recibir peticiones")
    print("  desde Excel VBA.")
    print()
    print("  Endpoints:")
    print("    GET  http://localhost:5050/health")
    print("    GET  http://localhost:5050/status")
    print("    POST http://localhost:5050/chat")
    print()
    print("  Presiona Ctrl+C para detener el servidor")
    print("=" * 50)
    
    app.run(host='127.0.0.1', port=5050, debug=False)
