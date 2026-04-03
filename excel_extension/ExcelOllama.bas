Attribute VB_Name = "OllamaIA"
'==============================================================================
' MÓDULO: OllamaIA para Excel - v4
' Usa texto plano con delimitadores (no JSON) para máxima compatibilidad
'==============================================================================

Private Const API_URL As String = "http://localhost:5050"
Private Const MAX_FILAS As Long = 10
Private Const BOTON_TAG As String = "OllamaIA_Boton"

'==============================================================================
' AUTO_OPEN: Crea botón en la barra al abrir el archivo
'==============================================================================
Public Sub Auto_Open()
    ' Quitar botón anterior si existe
    QuitarBoton
    
    ' Crear botón en la barra de herramientas
    On Error Resume Next
    Dim barra As Object
    Set barra = Application.CommandBars("Standard")
    
    Dim btn As Object
    Set btn = barra.Controls.Add(Type:=1) ' msoControlButton
    
    If Not btn Is Nothing Then
        btn.Caption = "Ollama IA"
        btn.Style = 3 ' msoButtonIconAndCaption
        btn.FaceId = 487 ' icono de robot/bombilla
        btn.Tag = BOTON_TAG
        btn.OnAction = "MostrarChatOllama"
        btn.TooltipText = "Abrir chat con Ollama IA"
    End If
    On Error GoTo 0
End Sub

'==============================================================================
' AUTO_CLOSE: Quita el botón al cerrar
'==============================================================================
Public Sub Auto_Close()
    QuitarBoton
End Sub

Private Sub QuitarBoton()
    On Error Resume Next
    Dim ctrl As Object
    For Each ctrl In Application.CommandBars("Standard").Controls
        If ctrl.Tag = BOTON_TAG Then ctrl.Delete
    Next ctrl
    On Error GoTo 0
End Sub

'==============================================================================
' PUNTO DE ENTRADA: Alt+F8 -> MostrarChatOllama
'==============================================================================
Public Sub MostrarChatOllama()
    If Not ServidorActivo() Then
        MsgBox "Servidor no disponible." & vbCrLf & _
               "Ejecuta 'run_excel_api.bat' primero.", vbCritical, "Excel + Ollama"
        Exit Sub
    End If
    
    Do
        Dim instruccion As String
        instruccion = InputBox( _
            "Escribe una instruccion:" & vbCrLf & vbCrLf & _
            "Ejemplos:" & vbCrLf & _
            "  - Suma 10 a la columna B" & vbCrLf & _
            "  - Crea una columna Observaciones" & vbCrLf & _
            "  - Ordena por columna A" & vbCrLf & _
            "  - Colorea en rojo celdas > 100" & vbCrLf & vbCrLf & _
            "'salir' para cerrar.", _
            "Excel + Ollama")
        
        If Len(Trim(instruccion)) = 0 Or LCase(Trim(instruccion)) = "salir" Then Exit Do
        
        Application.StatusBar = "Ollama pensando..."
        DoEvents
        
        ' Llamar al API (devuelve texto plano, no JSON)
        Dim respuestaRaw As String
        respuestaRaw = LlamarAPISimple(instruccion)
        
        Application.StatusBar = False
        
        ' Parsear respuesta con delimitadores
        Dim texto As String
        Dim codigo As String
        Call ParsearRespuesta(respuestaRaw, texto, codigo)
        
        If Len(codigo) > 5 Then
            ' Hay código VBA - mostrar y preguntar
            Dim pregunta As String
            pregunta = "Codigo VBA generado:" & vbCrLf & vbCrLf
            pregunta = pregunta & Left(codigo, 700)
            If Len(codigo) > 700 Then pregunta = pregunta & vbCrLf & "..."
            pregunta = pregunta & vbCrLf & vbCrLf & "Ejecutar?"
            
            If MsgBox(pregunta, vbYesNo + vbQuestion, "Ejecutar codigo?") = vbYes Then
                EjecutarCodigo codigo
            End If
        Else
            ' Solo texto
            If Len(texto) > 0 Then
                MsgBox Left(texto, 1000), vbInformation, "Respuesta de Ollama"
            Else
                MsgBox "No se recibio respuesta del servidor.", vbExclamation, "Excel + Ollama"
            End If
        End If
    Loop
End Sub

'==============================================================================
' LLAMAR AL API (texto plano, sin JSON)
'==============================================================================
Private Function LlamarAPISimple(instruccion As String) As String
    On Error GoTo ErrorHTTP
    
    Dim contexto As String
    contexto = LeerContextoHoja()
    
    ' Construir body JSON simple para el POST
    Dim body As String
    body = "{""instruccion"":""" & Esc(instruccion) & """,""contexto"":""" & Esc(contexto) & """}"
    
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "POST", API_URL & "/chat_simple", False
    http.setRequestHeader "Content-Type", "application/json"
    http.Send body
    
    If http.Status = 200 Then
        ' La respuesta es texto plano, no JSON!
        LlamarAPISimple = http.responseText
    Else
        LlamarAPISimple = "===TEXTO===" & vbCrLf & "Error HTTP " & http.Status & vbCrLf & "===FIN==="
    End If
    
    Set http = Nothing
    Exit Function
    
ErrorHTTP:
    LlamarAPISimple = "===TEXTO===" & vbCrLf & "Error: " & Err.Description & vbCrLf & "===FIN==="
End Function

'==============================================================================
' PARSEAR RESPUESTA (busca ===TEXTO=== y ===CODIGO===)
'==============================================================================
Private Sub ParsearRespuesta(raw As String, ByRef texto As String, ByRef codigo As String)
    texto = ""
    codigo = ""
    
    ' Buscar sección TEXTO
    Dim posTexto As Long
    posTexto = InStr(raw, "===TEXTO===")
    
    If posTexto = 0 Then
        ' Sin delimitadores, tratar todo como texto
        texto = raw
        Exit Sub
    End If
    
    ' Buscar sección CODIGO
    Dim posCodigo As Long
    posCodigo = InStr(raw, "===CODIGO===")
    
    ' Buscar FIN
    Dim posFin As Long
    posFin = InStr(raw, "===FIN===")
    If posFin = 0 Then posFin = Len(raw) + 1
    
    ' Extraer texto
    Dim textoInicio As Long
    textoInicio = posTexto + Len("===TEXTO===") + 1 ' +1 para el salto de línea
    
    If posCodigo > 0 Then
        texto = Mid(raw, textoInicio, posCodigo - textoInicio)
        ' Extraer código
        Dim codigoInicio As Long
        codigoInicio = posCodigo + Len("===CODIGO===") + 1
        codigo = Mid(raw, codigoInicio, posFin - codigoInicio)
    Else
        texto = Mid(raw, textoInicio, posFin - textoInicio)
    End If
    
    ' Limpiar
    texto = Trim(texto)
    codigo = Trim(codigo)
End Sub

'==============================================================================
' EJECUTAR CÓDIGO VBA
'==============================================================================
Private Sub EjecutarCodigo(codigo As String)
    On Error GoTo ErrorExec
    
    Application.ScreenUpdating = False
    
    Dim vbComp As Object
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(1)
    
    Dim src As String
    src = "Sub OllamaTemp()" & vbCrLf & codigo & vbCrLf & "End Sub"
    
    vbComp.CodeModule.AddFromString src
    Application.Run vbComp.Name & ".OllamaTemp"
    ThisWorkbook.VBProject.VBComponents.Remove vbComp
    
    Application.ScreenUpdating = True
    MsgBox "Ejecutado correctamente!", vbInformation, "Excel + Ollama"
    Exit Sub
    
ErrorExec:
    Dim errMsg As String
    errMsg = Err.Description
    On Error Resume Next
    If Not vbComp Is Nothing Then ThisWorkbook.VBProject.VBComponents.Remove vbComp
    Application.ScreenUpdating = True
    On Error GoTo 0
    
    MsgBox "Error: " & errMsg & vbCrLf & vbCrLf & _
           "Habilita en Excel:" & vbCrLf & _
           "Archivo > Opciones > Centro de confianza >" & vbCrLf & _
           "Config. macros > Marcar:" & vbCrLf & _
           "'Confiar en el acceso al modelo de objetos" & vbCrLf & _
           "de proyectos VBA'", vbCritical, "Excel + Ollama"
End Sub

'==============================================================================
' LEER CONTEXTO DE LA HOJA ACTIVA
'==============================================================================
Private Function LeerContextoHoja() As String
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim lr As Long, lc As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim s As String
    s = "Hoja:" & ws.Name & " Filas:" & lr & " Cols:" & lc & vbCrLf
    
    ' Encabezados
    Dim mc As Long
    mc = lc: If mc > 12 Then mc = 12
    
    s = s & "Encabezados: "
    Dim c As Long
    For c = 1 To mc
        s = s & CL(c) & "=" & CStr(ws.Cells(1, c).Value)
        If c < mc Then s = s & ", "
    Next c
    s = s & vbCrLf
    
    ' Datos
    Dim mr As Long
    mr = lr: If mr > MAX_FILAS Then mr = MAX_FILAS
    Dim mcd As Long
    mcd = lc: If mcd > 6 Then mcd = 6
    
    Dim r As Long
    For r = 2 To mr
        For c = 1 To mcd
            Dim v As String
            v = CStr(ws.Cells(r, c).Value)
            If Len(v) > 20 Then v = Left(v, 20) & ".."
            s = s & CL(c) & "=" & v
            If c < mcd Then s = s & "|"
        Next c
        s = s & vbCrLf
    Next r
    
    LeerContextoHoja = s
End Function

'==============================================================================
' UTILIDADES
'==============================================================================
Private Function ServidorActivo() As Boolean
    On Error GoTo No
    Dim h As Object
    Set h = CreateObject("MSXML2.XMLHTTP")
    h.Open "GET", API_URL & "/health", False
    h.Send
    ServidorActivo = (h.Status = 200)
    Set h = Nothing
    Exit Function
No:
    ServidorActivo = False
End Function

Private Function Esc(t As String) As String
    Dim r As String
    r = Replace(t, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, vbCrLf, "\n")
    r = Replace(r, vbCr, "\n")
    r = Replace(r, vbLf, "\n")
    r = Replace(r, vbTab, " ")
    Esc = r
End Function

Private Function CL(n As Long) As String
    Dim x As Long, s As String
    x = n
    Do While x > 0
        s = Chr(((x - 1) Mod 26) + 65) & s
        x = Int((x - 1) / 26)
    Loop
    CL = s
End Function
