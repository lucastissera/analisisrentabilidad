# Sistema de Análisis de Rentabilidad
### Backend Python (Flask) + Frontend HTML/JS

---

## Estructura del proyecto

```
.
├── app.py              ← Servidor Flask (API REST + rutas)
├── excel_engine.py     ← Motor Excel: exportación, importación, plantillas
├── requirements.txt   ← Dependencias Python
├── templates/
│   └── index.html      ← Frontend completo (HTML + CSS + JS)
└── data/               ← SQLite local (ignorada en git; p. ej. users.db)
```

---

## Instalación

### 1. Clonar / descomprimir el proyecto
```bash
cd <carpeta_del_repo>   # raíz clon del repositorio
```

### 2. Crear entorno virtual (recomendado)
```bash
python -m venv venv
source venv/bin/activate        # Linux / Mac
venv\Scripts\activate           # Windows
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

---

## Ejecución

```bash
python app.py
```

Abrí el navegador en: **http://localhost:5050**

---

## Funcionalidades

### Configuración inicial
- Nombre de empresa y período
- **Tipo de análisis**: Mensual (con carga mes a mes) o Anual
- Cantidad de sucursales (1 a 8) con nombres personalizados
- 6 métodos de prorrateo de gastos corporativos
- Carga de gastos corporativos por categoría

### Dashboard
- KPIs consolidados en tiempo real
- Gráficos de barras por sucursal (Margen Bruto y EBITDA)
- Cascada de resultados (Ventas → EBITDA)
- Selector de período + vista acumulada (YTD)

### Vistas
| Vista | Descripción |
|-------|-------------|
| Dashboard | KPIs, gráficos, waterfall |
| Comparativo | Tabla completa todas las sucursales |
| Semáforo | Verde/Amarillo/Rojo por KPI |
| Prorrateo | Comparativa de 6 métodos de asignación |
| Detalle Sucursal | Análisis individual por sucursal |
| Carga por período | Formulario mes a mes o año a año |

### Prorrateo de gastos corporativos
| Método | Criterio |
|--------|----------|
| % de Ventas | Proporcional a ventas netas |
| Nº Empleados | Por dotación de personal |
| M² del Local | Por espacio físico |
| Igualitario | Partes iguales |
| Transacciones | Por volumen de operaciones |
| Manual / Mixto | Porcentajes personalizados |

### Excel
- **Exportar**: genera un `.xlsx` con 5 hojas (Portada, Datos, Acumulado, Métricas, Corporativos)
- **Plantilla de importación**: descarga una plantilla en blanco lista para completar
- **Importar**: sube un `.xlsx` (exportado o plantilla completada) para cargar períodos

---

## API REST

| Método | Endpoint | Descripción |
|--------|----------|-------------|
| GET | `/` | Frontend HTML |
| POST | `/api/config` | Guardar configuración |
| GET | `/api/config` | Obtener configuración |
| POST | `/api/data/month` | Guardar datos de un período |
| DELETE | `/api/data/month/<period>` | Eliminar un período |
| GET | `/api/metrics` | Obtener métricas calculadas (todos los períodos) |
| GET | `/api/export/excel` | Descargar Excel con todos los datos |
| GET | `/api/template/excel` | Descargar plantilla de importación |
| POST | `/api/import/excel` | Importar desde Excel |

### Formato de período
- **Mensual**: `YYYY-MM` (ej: `2024-01` para Enero 2024)
- **Anual**: `YYYY` (ej: `2024`)

### Ejemplo de llamada a la API
```python
import requests

# Guardar datos de Enero 2024
requests.post('http://localhost:5050/api/data/month', json={
    "period": "2024-01",
    "branches": [
        {"ventasBrutas": 850000, "devoluciones": 15000, "otrosIngresos": 5000,
         "cmv": 340000, "moDirecta": 42500, "materiales": 12000,
         "alquiler": 28000, "sueldos": 95000, "servicios": 8500,
         "mantenimiento": 4200, "marketing": 12000, "admGeneral": 6500,
         "inventario": 120000, "equipamiento": 85000, "mobiliario": 40000,
         "empleados": 8, "m2": 180, "transacciones": 420},
        # ... más sucursales
    ]
})

# Obtener métricas
metrics = requests.get('http://localhost:5050/api/metrics').json()
print(metrics['2024-01']['totals'])
```

---

## Despliegue (Render y similares)

Debe escucharse en **`0.0.0.0`** y en el **puerto** que pone el host (`PORT` en Render), no fijo 5050.

- **Comando de inicio (Start Command)**, por ejemplo:  
  `gunicorn --bind 0.0.0.0:$PORT app:app`  
  o el `Procfile` de este repositorio (Gunicorn ya está en `requirements.txt`).

- **Build Command:** `pip install -r requirements.txt`

- **Root Directory:** dejá vacío si el repo en GitHub tiene `app.py` en la **raíz** (como ahora). Si usás otra ruta, indicala ahí.

- **Variable de entorno** `SECRET_KEY`: definila en Render (cadena larga y aleatoria) para que las sesiones de login sean fiables.

- Si al abrir el sitio ves *Not found* o error, revisá en **Logs** de Render: suele ser puerto no enlazado o el comando de arranque sin `gunicorn` / sin `--bind 0.0.0.0:$PORT`.

---

## Próximos pasos para escalar

1. **Base de datos**: reemplazar `_store` (dict en memoria) por SQLite o PostgreSQL
2. **Autenticación**: login por empresa / usuario
3. **Multi-empresa**: soporte para múltiples empresas en la misma instancia
4. **Gráficos de tendencia**: evolución mensual de KPIs con Chart.js o Recharts
5. **Exportar PDF**: reporte ejecutivo en PDF con weasyprint o reportlab
6. **Deploy**: Dockerfile incluido para despliegue en Railway, Render o AWS

---

## Notas de desarrollo (Cursor / VS Code)

El proyecto usa las siguientes tecnologías:
- **Backend**: Python 3.11+ / Flask 3.x
- **Excel**: openpyxl 3.x (creación/lectura), pandas (análisis)
- **Frontend**: HTML5 + CSS3 + JavaScript vanilla (sin frameworks)
- **Fuentes**: Google Fonts — Outfit + JetBrains Mono

Para desarrollo con recarga automática:
```bash
FLASK_ENV=development python app.py
```
