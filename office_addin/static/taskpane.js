/**
 * Excel IA Chat - Office Add-in Taskpane
 * Integración con Office.js para manipular Excel desde lenguaje natural
 */

// ============================================================================
// CONFIGURACIÓN
// ============================================================================

// URL del servidor backend (cambiar si se despliega en remoto)
const API_BASE = window.location.origin;

// Historial de conversación para contexto
let historial = [];

// Estado
let isProcessing = false;
let officeReady = false;

// ============================================================================
// INICIALIZACIÓN DE OFFICE.JS
// ============================================================================

Office.onReady(function (info) {
    if (info.host === Office.HostType.Excel) {
        officeReady = true;
        console.log("Office.js listo - Host: Excel");
        checkServerHealth();
    } else {
        updateBadge("No es Excel", "badge-error");
    }

    // Enter para enviar (Shift+Enter para nueva línea)
    document.getElementById("user-input").addEventListener("keydown", function (e) {
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            enviarMensaje();
        }
    });
});

// ============================================================================
// VERIFICAR CONEXIÓN CON SERVIDOR
// ============================================================================

async function checkServerHealth() {
    try {
        const res = await fetch(API_BASE + "/api/health");
        const data = await res.json();

        if (data.ia_disponible) {
            const provider = data.ia_provider;
            const label = provider === "claude" ? "Claude" : "Gemini";
            const cls = provider === "claude" ? "badge-claude" : "badge-gemini";
            updateBadge(label, cls);
        } else {
            updateBadge("Sin IA", "badge-error");
        }
    } catch (e) {
        updateBadge("Sin conexión", "badge-error");
    }
}

function updateBadge(text, className) {
    const badge = document.getElementById("status-badge");
    badge.textContent = text;
    badge.className = "badge " + className;
}

// ============================================================================
// OBTENER CONTEXTO DE EXCEL (Office.js)
// ============================================================================

async function getExcelContext() {
    if (!officeReady) return "No hay conexión con Excel";

    try {
        return await Excel.run(async function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");

            var usedRange = sheet.getUsedRange();
            usedRange.load(["address", "values", "rowCount", "columnCount"]);

            // Obtener selección actual
            var selection = context.workbook.getSelectedRange();
            selection.load(["address", "values", "rowCount", "columnCount"]);

            await context.sync();

            var headers = usedRange.values[0] || [];
            var sampleRows = usedRange.values.slice(1, 6);

            var info = "Hoja: " + sheet.name + "\n";
            info += "Rango total: " + usedRange.address + "\n";
            info += "Filas: " + usedRange.rowCount + " | Columnas: " + usedRange.columnCount + "\n";
            info += "Encabezados: " + headers.join(", ") + "\n";
            info += "Muestra de datos:\n";

            for (var i = 0; i < sampleRows.length; i++) {
                info += "  Fila " + (i + 2) + ": " + sampleRows[i].join(" | ") + "\n";
            }

            // Incluir información de la selección
            info += "\nSELECCION ACTUAL: " + selection.address + "\n";
            info += "Tamanio seleccion: " + selection.rowCount + " filas x " + selection.columnCount + " columnas\n";

            if (selection.rowCount <= 50 && selection.columnCount <= 20) {
                info += "Datos seleccionados:\n";
                for (var j = 0; j < selection.values.length; j++) {
                    info += "  " + selection.values[j].join(" | ") + "\n";
                }
            } else {
                // Si la selección es muy grande, solo mostrar muestra
                info += "Datos seleccionados (muestra, primeras 10 filas):\n";
                var maxRows = Math.min(10, selection.values.length);
                for (var k = 0; k < maxRows; k++) {
                    info += "  " + selection.values[k].join(" | ") + "\n";
                }
                info += "  ... (" + selection.rowCount + " filas en total)\n";
            }

            return info;
        });
    } catch (e) {
        return "Error obteniendo contexto: " + e.message;
    }
}

// ============================================================================
// EJECUTAR CÓDIGO OFFICE.JS GENERADO POR LA IA
// ============================================================================

async function ejecutarCodigo(codigoJS) {
    if (!officeReady) {
        addStatusMessage("No hay conexión con Excel", true);
        return false;
    }

    try {
        await Excel.run(async function (context) {
            var sheet = context.workbook.worksheets.getActiveWorksheet();
            sheet.load("name");
            await context.sync();

            // Ejecutar el código generado
            var asyncFn = new Function(
                "context", "sheet", "Excel",
                "return (async function() {\n" + codigoJS + "\n})();"
            );

            await asyncFn(context, sheet, Excel);
        });

        addStatusMessage("Ejecutado correctamente", false);
        return true;
    } catch (e) {
        addStatusMessage("Error: " + e.message, true);
        return false;
    }
}

// ============================================================================
// ENVIAR MENSAJE
// ============================================================================

async function enviarMensaje() {
    var input = document.getElementById("user-input");
    var texto = input.value.trim();

    if (!texto || isProcessing) return;

    isProcessing = true;
    input.value = "";
    document.getElementById("send-btn").disabled = true;

    // Mostrar mensaje del usuario
    addMessage(texto, "user");

    // Mostrar indicador de escritura
    var typingId = showTyping();

    try {
        // Obtener contexto de Excel
        var contexto = await getExcelContext();

        // Llamar al servidor
        var res = await fetch(API_BASE + "/api/chat", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                instruccion: texto,
                contexto: contexto,
                historial: historial
            })
        });

        var data = await res.json();

        // Quitar indicador de escritura
        removeTyping(typingId);

        if (data.error) {
            addMessage("Error: " + data.error, "bot");
        } else {
            // Guardar en historial
            historial.push({ rol: "usuario", texto: texto });
            historial.push({ rol: "asistente", texto: data.respuesta });

            // Limitar historial
            if (historial.length > 20) {
                historial = historial.slice(-20);
            }

            // Mostrar respuesta
            addBotResponse(data.respuesta, data.codigo_js, data.tiene_codigo);

            // Auto-ejecutar si está habilitado
            if (data.tiene_codigo && document.getElementById("auto-execute").checked) {
                await ejecutarCodigo(data.codigo_js);
            }
        }
    } catch (e) {
        removeTyping(typingId);
        addMessage("Error de conexión: " + e.message, "bot");
    }

    isProcessing = false;
    document.getElementById("send-btn").disabled = false;
    input.focus();
}

// ============================================================================
// UI - MENSAJES
// ============================================================================

function addMessage(text, type) {
    var container = document.getElementById("chat-container");
    var div = document.createElement("div");
    div.className = "message " + (type === "user" ? "user-message" : "bot-message");

    var content = document.createElement("div");
    content.className = "message-content";
    content.textContent = text;

    div.appendChild(content);
    container.appendChild(div);
    scrollToBottom();
}

function addBotResponse(respuesta, codigoJS, tieneCodigo) {
    var container = document.getElementById("chat-container");
    var div = document.createElement("div");
    div.className = "message bot-message";

    var content = document.createElement("div");
    content.className = "message-content";

    // Limpiar la respuesta: quitar el bloque de código del texto visible
    var textoLimpio = respuesta
        .replace(/```javascript[\s\S]*?```/gi, "")
        .replace(/```js[\s\S]*?```/gi, "")
        .trim();

    // Convertir markdown básico a HTML
    textoLimpio = textoLimpio
        .replace(/\*\*(.*?)\*\*/g, "<strong>$1</strong>")
        .replace(/\n/g, "<br/>");

    content.innerHTML = textoLimpio;

    // Mostrar bloque de código si existe
    if (tieneCodigo && codigoJS) {
        var codeBlock = document.createElement("div");
        codeBlock.className = "code-block";
        codeBlock.textContent = codigoJS;
        content.appendChild(codeBlock);

        // Botón de ejecutar (si no hay auto-ejecución)
        if (!document.getElementById("auto-execute").checked) {
            var actions = document.createElement("div");
            actions.className = "message-actions";

            var execBtn = document.createElement("button");
            execBtn.className = "action-btn execute-btn";
            execBtn.textContent = "Ejecutar";
            execBtn.onclick = function () {
                ejecutarCodigo(codigoJS);
                execBtn.disabled = true;
                execBtn.textContent = "Ejecutado";
            };
            actions.appendChild(execBtn);

            content.appendChild(actions);
        }
    }

    div.appendChild(content);
    container.appendChild(div);
    scrollToBottom();
}

function addStatusMessage(text, isError) {
    var container = document.getElementById("chat-container");
    var div = document.createElement("div");
    div.className = "status-message " + (isError ? "status-error" : "status-success");
    div.textContent = isError ? "✕ " + text : "✓ " + text;
    container.appendChild(div);
    scrollToBottom();
}

function showTyping() {
    var container = document.getElementById("chat-container");
    var div = document.createElement("div");
    var id = "typing-" + Date.now();
    div.id = id;
    div.className = "typing-indicator";
    div.innerHTML = "<span></span><span></span><span></span>";
    container.appendChild(div);
    scrollToBottom();
    return id;
}

function removeTyping(id) {
    var el = document.getElementById(id);
    if (el) el.remove();
}

function scrollToBottom() {
    var container = document.getElementById("chat-container");
    container.scrollTop = container.scrollHeight;
}
