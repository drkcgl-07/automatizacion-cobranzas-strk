# EVIDENCIA DE COLABORACIÓN DEL EQUIPO - SISTEMA ARPC

## 🧑‍💻 INTEGRANTES DEL EQUIPO:
| 1 | **Carlos García** | Líder del Proyecto y Desarrollador Principal | Arquitectura del sistema, interfaz de usuario, integración total |
| 2 | **Elisa Cunya** | Analista de Datos | Validaciones, estadísticas, gestión de documentos |
| 3 | **Josemir Poma** | Especialista en Algoritmos | Cálculos de tramos, proyecciones, lógica de negocio |
| 4 | **Josie Mamani** | Especialista en Reportes | Exportación Excel, formatos, documentación |

## 📊 DISTRIBUCIÓN DE CÓDIGO EN `main.py` (628 líneas totales):

### **JOSEMIR POMA** (120 líneas ≈ 19%)
- **Líneas 42-176**: Clase `DocumentoSAP` completa
  - Método `_calcular_tramo()` (líneas 87-101)
  - Método `_calcular_estatus()` (líneas 103-108) 
  - Método `_calcular_proyeccion()` (líneas 110-176)
  - Algoritmos de cálculo de semanas y tramos

### **ELISA CUNYA** (63 líneas ≈ 10%)
- **Líneas 178-210**: Clase `GestorDocumentos` completa
  - Gestión de colección de documentos
  - Método `obtener_estadisticas()`
  - Filtrado y validación
- **Líneas 212-233**: Clase `GeneradorTablasDinamicas` (parte inicial)

### **JOSIE MAMANI** (85 líneas ≈ 14%)
- **Líneas 380-465**: Método `exportar_reporte_completo()` completo
  - Exportación a Excel con 3 hojas
  - Formato y estructura de reportes
  - Generación de archivos Excel profesionales

### **CARLOS GARCÍA** (360 líneas ≈ 57%)
- **Líneas 1-41**: Excepciones e interfaces
- **Líneas 234-258**: Resto de `GeneradorTablasDinamicas`
- **Líneas 260-379**: Clase `ProcesadorARPC` (excepto exportación)
- **Líneas 467-492**: Clase `SistemaARPC` (interfaz)
- **Líneas 494-540**: Función `main()` e integración total
- Todo el sistema de archivos y selección
- Interfaz de usuario completa (tkinter + consola)

## 🤝 PROCESO DE COLABORACIÓN:

### **Fase 1: Análisis (Semana 1)**
- Reunión grupal: Análisis del proceso manual
- Josemir investigó fórmulas de cálculo
- Elisa analizó estructura de datos SAP
- Carlos diseñó arquitectura del sistema

### **Fase 2: Desarrollo (Semanas 2-3)**
- Josemir desarrolló algoritmos de tramos
- Elisa implementó gestión de documentos
- Josie trabajó en exportación Excel
- Carlos integró todos los módulos

### **Fase 3: Pruebas (Semana 4)**
- Pruebas con datos reales: Todos
- Validación de cálculos: Josemir + Elisa
- Prueba de exportación: Josie
- Prueba de interfaz: Carlos

## 📎 ARCHIVOS ADJUNTOS:
1. Capturas de pantalla de reuniones (Zoom/Meet)
2. Documento de requisitos inicial
3. Diagramas de flujo del proceso
4. Ejemplos de datos de prueba
