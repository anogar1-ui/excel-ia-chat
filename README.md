# 📊 Gestor de Excel con Streamlit

Aplicación web interactiva para cargar, visualizar, editar, analizar y exportar archivos Excel.

## 🚀 Instalación

```bash
# Instalar dependencias
pip install -r requirements.txt
```

## ▶️ Ejecutar la Aplicación

```bash
streamlit run app.py
```

La aplicación se abrirá automáticamente en tu navegador en `http://localhost:8501`

## 📋 Funcionalidades

| Función | Descripción |
|---------|-------------|
| **Cargar Excel** | Sube archivos .xlsx desde la barra lateral |
| **Editar Datos** | Modifica celdas directamente en la tabla |
| **Fórmulas** | Aplica operaciones matemáticas entre columnas |
| **Gráficos** | Genera visualizaciones de barras, líneas o dispersión |
| **Chat IA** | Escribe comandos en lenguaje natural |
| **Exportar** | Descarga el archivo modificado |

## 💬 Comandos del Chat IA

- `suma 10 a Precio` - Suma 10 a toda la columna Precio
- `multiplica Cantidad por 2` - Multiplica los valores
- `filtra donde Precio > 100` - Filtra filas
- `ordena por Nombre` - Ordena la tabla
- `gráfico de barras de Ventas` - Genera un gráfico

## 📁 Estructura del Proyecto

```
excellll/
├── app.py              # Aplicación principal
├── requirements.txt    # Dependencias
└── README.md          # Esta documentación
```
