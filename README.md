# 📊 Generador de Reportes de Inscripciones a Cursos

Una aplicación web desarrollada con **Streamlit** que genera automáticamente reportes detallados de inscripciones a cursos a partir de archivos Excel. La aplicación procesa los datos, genera visualizaciones y crea un documento Word profesional listo para descargar.

## 🌟 Características Principales

- ✅ **Interfaz web intuitiva** y fácil de usar
- 📁 **Carga de archivos Excel** desde el computador
- 📊 **Visualizaciones automáticas** con gráficos de barras
- 📄 **Generación de reportes Word** profesionales
- 📈 **Análisis estadístico** completo de los datos
- 🔍 **Validación automática** de la estructura de datos
- ⬇️ **Descarga directa** del reporte generado

## 🚀 Demo

![Captura de pantalla de la aplicación](https://via.placeholder.com/800x400/4CAF50/FFFFFF?text=Generador+de+Reportes)

## 📋 Requisitos del Sistema

- Python 3.8 o superior
- Conexión a internet (para la primera instalación)

## 🛠️ Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/tu-usuario/generador-reportes-inscripciones.git
cd generador-reportes-inscripciones
```

### 2. Crear un entorno virtual (recomendado)
```bash
# Windows
python -m venv venv
venv\Scripts\activate

# macOS/Linux
python3 -m venv venv
source venv/bin/activate
```

### 3. Instalar las dependencias
```bash
pip install -r requirements.txt
```

## 🎯 Uso de la Aplicación

### 1. Ejecutar la aplicación
```bash
streamlit run app.py
```

### 2. Abrir en el navegador
La aplicación se abrirá automáticamente en `http://localhost:8501`

### 3. Cargar archivo Excel
- Usa el panel lateral para cargar tu archivo Excel
- El archivo debe contener las columnas requeridas (ver estructura de datos)

### 4. Generar reporte
- Revisa las estadísticas y visualizaciones generadas
- Haz clic en "Generar Reporte Word"
- Descarga el reporte completo

## 📊 Estructura de Datos Requerida

Tu archivo Excel debe contener las siguientes columnas:

| Columna | Descripción | Ejemplo |
|---------|-------------|---------|
| `Nombre y apellidos completos` | Nombre completo del inscrito | Juan Pérez García |
| `Hora de inicio` | Fecha y hora de inscripción | 2024-01-15 09:30:00 |
| `Curso de interés` | Nombre del curso | Python Básico |
| `Correo de contacto` | Email del inscrito | juan.perez@email.com |

### Ejemplo de datos:
```
Nombre y apellidos completos | Hora de inicio      | Curso de interés | Correo de contacto
Juan Pérez García           | 2024-01-15 09:30:00 | Python Básico    | juan.perez@email.com
María López Rodríguez       | 2024-01-16 14:20:00 | Excel Avanzado   | maria.lopez@email.com
Carlos Martín Sánchez       | 2024-01-17 11:45:00 | Python Básico    | carlos.martin@email.com
```

## 📑 Contenido del Reporte Generado

El reporte Word incluye:

### 1. **Portada**
- Título del reporte
- Fecha de elaboración
- Información del elaborador

### 2. **Resumen General**
- Total de personas inscritas (valores únicos)
- Fecha de inicio y finalización de inscripciones
- Estadísticas generales

### 3. **Visualización**
- Gráfico de barras de inscripciones por curso
- Datos numéricos en cada barra

### 4. **Detalle por Curso**
- Tablas organizadas por curso
- Lista de inscritos con nombres y correos
- Contador de inscritos por curso

## 🔧 Dependencias

```
streamlit==1.28.0
pandas==2.0.3
matplotlib==3.7.2
python-docx==0.8.11
numpy==1.24.3
openpyxl==3.1.2
```

## 📁 Estructura del Proyecto

```
generador-reportes-inscripciones/
│
├── app.py                 # Aplicación principal de Streamlit
├── requirements.txt       # Dependencias del proyecto
├── README.md             # Documentación
└── ejemplos/             # Archivos de ejemplo (opcional)
    └── datos_ejemplo.xlsx
```

## ⚠️ Solución de Problemas

### Error: "Faltan columnas requeridas"
- Verifica que tu archivo Excel contenga exactamente las columnas mencionadas
- Asegúrate de que los nombres de las columnas coincidan exactamente

### Error: "Error al procesar el archivo"
- Verifica que el archivo no esté corrupto
- Asegúrate de que el archivo sea un Excel válido (.xlsx o .xls)
- Revisa que las fechas estén en formato correcto

### La aplicación no se abre
- Verifica que todas las dependencias estén instaladas
- Asegúrate de estar en el directorio correcto
- Verifica que el puerto 8501 no esté ocupado

## 🎨 Personalización

### Cambiar el autor del reporte
Modifica la línea 93 en `app.py`:
```python
parrafo_autor = documento.add_paragraph('Elaborado por: Tu Nombre Aquí')
```

### Cambiar colores del gráfico
Modifica la línea 52 en `app.py`:
```python
bars = ax.bar(..., color='tu_color_aqui', edgecolor='tu_borde_aqui')
```

### Personalizar el título
Modifica la configuración de la página en `app.py`:
```python
st.set_page_config(
    page_title="Tu Título Personalizado",
    page_icon="🎯"
)
```



## 📊 Estadísticas del Proyecto

- ⭐ **Lenguaje principal**: Python
- 📦 **Dependencias**: 6 principales
- 🔧 **Compatibilidad**: Windows, macOS, Linux
- 📱 **Interfaz**: Web responsiva

---



¡Gracias por usar el Generador de Reportes de Inscripciones! 🚀
